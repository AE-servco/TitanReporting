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

def format_employee_list(employee_response):
    # input can be either technician response or employee response
    formatted = {}
    for employee in employee_response:
        formatted[employee['id']] = employee['name']
        formatted[employee['userId']] = employee['name']
    return formatted

def get_all_employee_ids(client: ServiceTitanClient):
    tech_url = client.build_url("settings", "technicians")
    techs = format_employee_list(client.get_all(tech_url))
    emp_url = client.build_url("settings", "employees")
    office = format_employee_list(client.get_all(emp_url))
    return techs | office

def fetch_jobs(
    start_date: _dt.date,
    end_date: _dt.date,
    _client: ServiceTitanClient,
    job_id: str = None,
    status_filters: List = [],
) -> List[Dict[str, Any]]:
    """
    Retrieve all jobs created between `start_date` and `end_date`,
    converting the local date boundaries into UTC timestamps. If
    job_num specified, just fetches that job.
    """

    tenant = _client.tenant or "{tenant}"
    base_path = f"jpm/v2/tenant/{tenant}/jobs"
    jobs: List[Dict[str, Any]] = []

    created_after = _client.start_of_day_utc_string(start_date)
    created_before = _client.end_of_day_utc_string(end_date)

    # If job_id specified, only return that job
    if job_id:
        params = {
            "ids": job_id,
        }
        try:
            resp = _client.get(base_path, params=params)
        except Exception:
            return []
        if not isinstance(resp, dict):
            return []
        page_data: Iterable[Dict[str, Any]] = resp.get("data") or []
        return page_data
    params = {
                "createdOnOrAfter": created_after,
                "createdBefore": created_before,
            }
    if _client.app_guid:
        params["externalDataApplicationGuid"] = _client.app_guid
    if status_filters:
        for status in status_filters:
            params["jobStatus"] = status
            jobs.extend(_client.get_all(base_path, params=params))
    else:
        jobs = _client.get_all(base_path, params=params)
    return jobs

def fetch_invoices(
    ids: List,
    _client: ServiceTitanClient,
) -> List[Dict[str, Any]]:
    """
    Retrieve all invoices given a list of ids
    """
    if type(ids[0]) != str:
        ids = [str(id) for id in ids]
    base_path = _client.build_url('accounting', 'invoices')

    invoices = _client.get_all_id_filter(base_path, ids=ids)
    return invoices

def fetch_payments(
    invoice_ids: List,
    _client: ServiceTitanClient,
) -> List[Dict[str, Any]]:
    """
    Retrieve all invoices given a list of ids
    """
    if type(invoice_ids[0]) != str:
        invoice_ids = [str(id) for id in invoice_ids]
    base_path = _client.build_url('accounting', 'payments')

    invoices = _client.get_all_id_filter(base_path, ids=invoice_ids, id_filter_name='appliedToInvoiceIds')
    return invoices

def get_job_external_data(job_id, client, application_guid):
    url = client.build_url('jpm', 'jobs', resource_id=job_id)
    params = {"externalDataApplicationGuid": application_guid}
    job_data = client.get(url, params=params)
    external_entries = job_data.get("externalData", [])
    for entry in external_entries:
        if entry.get("key") == "docchecks":
            try:
                return json.loads(entry["value"])
            except Exception:
                return {}
    return {}

def get_doc_check_criteria():
    checks = {
        'pb': 'Before Photo',
        'pa': 'After Photo',
        'pr': 'Receipt Photo',
        'qd': 'Quote Description',
        'qs': 'Quote Signed',
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
            jobs = fetch_jobs(start_date, end_date, client, custom_job_id)
        else:
            jobs = fetch_jobs(start_date, end_date, client, status_filters=job_status_filter)
            if filter_unsucessful:
                jobs = filter_out_unsuccessful_jobs(jobs, client)
        st.session_state.jobs = jobs
        st.session_state.current_index = 0
        st.session_state.prefetched = {}
        st.session_state.prefetch_futures = {}

def get_invoice_ids(job_response):
    return [str(job['invoiceId']) for job in job_response]

def format_job(job, client: ServiceTitanClient):
    # if 116255355 in job['tagTypeIds'] or job['jobStatus'] == 'Canceled': # Unsuccessful or cancelled 
    #     return None
    formatted = {}
    if job['soldById'] is not None:
        formatted['Sold By'] = job['soldById']
    else:
        url = client.build_url("dispatch", "appointment-assignments")
        appts = client.get_all(url, params={'jobId': job['id']})
        if appts:
            formatted['Sold By'] = ', '.join([str(appt['technicianId']) for appt in appts])
        else:
            formatted['Sold By'] = -1
    formatted['Created Date'] = client.st_date_to_local(job['createdOn'], fmt="%m/%d/%Y")
    formatted['Completion Date'] = client.st_date_to_local(job['completedOn'], fmt="%m/%d/%Y") if job['completedOn'] is not None else "None"
    formatted['Job #'] = job['jobNumber'] if job['jobNumber'] is not None else "None"
    formatted['Status'] = job['jobStatus'] if job['jobStatus'] is not None else "None"
    formatted['invoiceId'] = job['invoiceId'] if job['invoiceId'] is not None else "None"
    formatted['externalData'] = job['externalData'][0] if job['externalData'] is not None and job['externalData'] != [] else "None"
    return formatted

def format_invoice(invoice):
    formatted = {}
    formatted['Suburb'] = invoice['customerAddress']['city']
    formatted['Jobs Subtotal'] = invoice['subTotal']
    formatted['Invoice Balance'] = invoice['balance']
    formatted['Payments'] = round(float(invoice['total']) - float(invoice['balance']),2)
    formatted['invoiceId'] = invoice['id']
    return formatted

def format_payment(payment):
    output = []
    for invoice in payment['appliedTo']:
        formatted = {}
        formatted['invoiceId'] = invoice['appliedTo']
        formatted['Payment Types'] = payment['type']
        output.append(formatted)
    return output