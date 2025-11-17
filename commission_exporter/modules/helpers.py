from __future__ import annotations

import datetime as _dt
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json

import streamlit as st
import streamlit_authenticator as stauth
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

from servicetitan_api_client import ServiceTitanClient
import modules.google_store as gs

import modules.data_formatting as format
import modules.data_fetching as fetching
import modules.lookup_tables as lookup


def flatten_list(nested_list):
    return [item for sublist in nested_list for item in sublist]

def get_client(tenant) -> ServiceTitanClient:
    @st.cache_resource(show_spinner=False)
    def _create_client(tenant) -> ServiceTitanClient:
            client = ServiceTitanClient(
                app_key=gs.get_secret("ST_app_key_tester"), 
                app_guid=gs.get_secret("ST_servco_integrations_guid"), 
                tenant=gs.get_secret(f"ST_tenant_id_{tenant}"), 
                client_id=gs.get_secret(f"ST_client_id_{tenant}"), 
                client_secret=gs.get_secret(f"ST_client_secret_{tenant}"), 
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


def filter_out_unsuccessful_jobs(jobs, client: ServiceTitanClient):
    unsuccessful_tags = [tag.get("id") for tag in fetching.fetch_tag_types(client) if "Unsuccessful" in tag.get("name")]
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