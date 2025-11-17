from __future__ import annotations

import datetime as _dt
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json
from google.cloud import secretmanager

import streamlit as st
import streamlit_authenticator as stauth
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

from servicetitan_api_client import ServiceTitanClient
import modules.google_store as gs

import modules.data_formatting as format
import modules.data_fetching as fetching

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

def flatten_list(nested_list):
    return [item for sublist in nested_list for item in sublist]

def get_secret(secret_id, project_id="servco1", version_id="latest"):
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/{secret_id}/versions/{version_id}"
    response = client.access_secret_version(request={"name": name})
    secret_payload = response.payload.data.decode("UTF-8")
    return secret_payload

def state_codes():
    codes = {
        'NSW_old': 'alphabravo',
        'VIC_old': 'victortango',
        'QLD_old': 'echozulu',
        'NSW': 'foxtrotwhiskey',
        'WA': 'sierradelta',
        'QLD': 'bravogolf',
    }
    return codes

def get_client(tenant) -> ServiceTitanClient:
    @st.cache_resource(show_spinner=False)
    def _create_client(tenant) -> ServiceTitanClient:
            # state_code = state_codes()[state]
            client = ServiceTitanClient(
                app_key=get_secret("ST_app_key_tester"), 
                app_guid=get_secret("ST_servco_integrations_guid"), 
                tenant=get_secret(f"ST_tenant_id_{tenant}"), 
                client_id=get_secret(f"ST_client_id_{tenant}"), 
                client_secret=get_secret(f"ST_client_secret_{tenant}"), 
                environment="production"
            )
            return client
    return _create_client(tenant)

# def get_sales_and_installer_codes(roles_reponse):
#     role_codes = {
#         's': [],
#         'i': []
#     }
#     for role in roles_reponse:
#         if role['name'] == 'Technician - Sales':
#             role_codes['s'].append(role['id'])
#         if role['name'] == 'Technician - Installer':
#             role_codes['i'].append(role['id'])
#     return role_codes

# def format_employee_list(employee_response, role_codes=None):
#     # input can be either technician response or employee response

#     def test_role(emp_roles, test_roles):
#         return bool(set(emp_roles) & set(test_roles))

#     formatted = {}
#     sales = set()
#     for employee in employee_response:
#         if role_codes:
#             test_s = test_role(employee['roleIds'], role_codes['s'])
#             test_i = test_role(employee['roleIds'], role_codes['i'])
#             formatted[employee['id']] = {'name': employee['name'], 'sales': test_s, 'installer': test_i}
#             formatted[employee['userId']] = {'name': employee['name'], 'sales': test_s, 'installer': test_i}
#             if test_s:
#                 sales.add(employee['id'])
#                 sales.add(employee['userId'])
#         else:   
#             formatted[employee['id']] = {'name': employee['name'], 'sales': False}
#             formatted[employee['userId']] = {'name': employee['name'], 'sales': False}
#     return formatted, sales


def get_all_employee_ids(client: ServiceTitanClient):
    roles_url = client.build_url("settings", "user-roles")
    sales_codes = get_sales_codes(client.get_all(roles_url))
    tech_url = client.build_url("settings", "technicians")
    techs, techs_sales = format.format_employee_list(client.get_all(tech_url), sales_codes)
    emp_url = client.build_url("settings", "employees")
    office, _ = format.format_employee_list(client.get_all(emp_url))
    return techs | office, techs_sales

def get_sales_codes(roles_reponse):
    sales_codes = set()
    for role in roles_reponse:
        if role['name'] == 'Technician - Sales':
            sales_codes.add(role['id'])
    return sales_codes


def get_doc_check_criteria():
    checks = {
        'pb': 'Before Photo',
        'pa': 'After Photo',
        'pr': 'Receipt Photo',
        'qd': 'Quote Description',
        'qs': 'Quote Description',
        'qe': 'Quote Emailed',
        'id': 'Invoice Description',
        'is': 'Invoice Signed',
        'ie': 'Invoice Emailed',
        '5s': '5 Star Review',
    }
    return checks

def get_tag_types(client: ServiceTitanClient):
    url = client.build_url('settings', 'tag-types')
    return client.get_all(url)

def filter_out_unsuccessful_jobs(jobs, client: ServiceTitanClient):
    unsuccessful_tags = [tag.get("id") for tag in get_tag_types(client) if "Unsuccessful" in tag.get("name")]
    return [job for job in jobs if unsuccessful_tags[0] not in job.get("tagTypeIds")]


def fetch_jobs_button_call(tenant_filter, start_date, end_date, job_status_filter, filter_unsucessful, custom_job_id=None):
    with st.spinner("Retrieving jobs..."):
        tenant_filter = tenant_filter.split(" ")[0].lower()
        st.session_state.current_tenant = tenant_filter
        client = st.session_state.clients.get(tenant_filter)
        if custom_job_id:
            jobs = fetching.fetch_jobs(start_date, end_date, client, custom_job_id)
        else:
            jobs = fetching.fetch_jobs(start_date, end_date, client, status_filters=job_status_filter)
            if filter_unsucessful:
                jobs = filter_out_unsuccessful_jobs(jobs, client)
        st.session_state.jobs = jobs
        st.session_state.current_index = 0
        st.session_state.prefetched = {}
        st.session_state.prefetch_futures = {}

def get_invoice_ids(job_response):
    return [str(job['invoiceId']) for job in job_response]


def categorise_job(job):
    status = job['status']
    day = job['created_dt'].weekday()
    balance = float(job['balance'])
    if day <5: # weekdays
        if status == 'Completed' and balance == 0:
            return 'wk_complete_paid'
        if status == 'Completed' and balance != 0:
            return 'wk_complete_unpaid'
        if status == 'Hold':
            return 'wk_wo'
            # return 'wk_hold'
        if status == 'In Progress':
            return 'wk_wo'
            # return 'wk_progress'
        if status == 'Scheduled':
            return 'wk_wo'
            # return 'wk_scheduled'
        if job['unsuccessful']:
            return 'wk_unsucessful'
        return 'wk_uncategorised'
    if day >=5: # weekdays
        if status == 'Completed' and balance == 0:
            return 'wkend_complete_paid'
        if status == 'Completed' and balance != 0:
            return 'wkend_complete_unpaid'
        if status == 'Hold':
            return 'wkend_wo'
            # return 'wkend_hold'
        if status == 'In Progress':
            return 'wkend_wo'
            # return 'wkend_progress'
        if status == 'Scheduled':
            return 'wkend_wo'
            # return 'wkend_scheduled'
        if job['unsuccessful']:
            return 'wkend_unsucessful'
        return 'wkend_uncategorised'