from __future__ import annotations

import datetime as _dt
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json
from google.cloud import secretmanager
from supabase import create_client, Client

import streamlit as st

from servicetitan_api_client import ServiceTitanClient
import modules.google_store as gs
import modules.formatting as format
import modules.fetching as fetch

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

def flatten_list(nested_list):
    return [item for sublist in nested_list for item in sublist]

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

def get_secret(secret_id, project_id="prestigious-gcp", version_id="latest"):
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/{secret_id}/versions/{version_id}"
    response = client.access_secret_version(request={"name": name})
    secret_payload = response.payload.data.decode("UTF-8")
    return secret_payload


def get_supabase() -> Client:
    url: str = get_secret("supabase_url")
    key: str = get_secret("supabase_secret_key")
    # key: str = get_secret("supabase_key")
    return create_client(url, key)

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

def get_all_employee_ids(client: ServiceTitanClient):
    tech_url = client.build_url("settings", "technicians")
    techs = format.format_employee_list(client.get_all(tech_url))
    emp_url = client.build_url("settings", "employees")
    office = format.format_employee_list(client.get_all(emp_url))
    return techs | office

def filter_image_attachments(attachments: List[Dict[str, Any]]) -> List[Tuple[str, int]]:
    """Filter attachments for supported image types.

    Returns a list of ``(filename, id)`` pairs for attachments whose
    filename ends with a recognised image extension.
    """
    results: List[Tuple[str, int]] = []
    for att in attachments:
        name = att.get("fileName") or att.get("filename") or att.get("name")
        att_id = att.get("id")
        if not name or att_id is None:
            continue
        ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
        if f".{ext}" in IMAGE_EXTENSIONS:
            try:
                results.append((name, int(att_id)))
            except Exception:
                pass
    return results

def group_attachments_by_type(
    attachments: List[Dict[str, Any]],
    extension_map: Optional[Dict[str, Set[str]]] = None,
) -> Dict[str, List[int]]:
    """
    Group attachments into categories based on file extension.

    :param attachments: A list of attachment dicts; each should have at least
                        'id' and 'fileName'.
    :param extension_map: Optional mapping of category names to sets of file
                          extensions (including the dot). If omitted, images
                          go under 'imgs' and PDFs under 'pdfs'.
    :return: A dict where each key is a category name and the value is a list
             of attachment IDs that match that category.

    Example of custom map:
    custom_map = {
            "imgs": {".jpg", ".png"},
            "pdfs": {".pdf"},
            "videos": {".mp4", ".mov"},
            "spreadsheets": {".xls", ".xlsx", ".csv"},
    }
    """
    if extension_map is None:
        extension_map = {
            "imgs": set(IMAGE_EXTENSIONS),
            "pdfs": {".pdf"},
        }

    result = {category: [] for category in extension_map}
    for att in attachments:
        file_name = att.get("fileName")
        att_id = att.get("id")
        file_date = att.get("createdOn")
        file_by = att.get("createdById")
        if not file_name or att_id is None:
            continue

        # Extract the lower‑cased extension with a leading dot
        ext = file_name.lower().rsplit(".", 1)[-1] if "." in file_name else ""
        ext_with_dot = f".{ext}" if ext else ""
        for category, exts in extension_map.items():
            if ext_with_dot in exts:
                result[category].append((file_name, int(att_id), file_date, file_by))
                break  # stop after the first matching category
    return result

def pre_fill_quote_signed_check(pdfs):
    if pdfs:
        for pdf in pdfs:
            fname = pdf.get('file_name').lower()
            if "estimate" in fname and "signed" in fname:
                return 1
    return 0

def pre_fill_invoice_signed_check(pdfs):
    if pdfs:
        for pdf in pdfs:
            fname = pdf.get('file_name').lower()
            if "invoice" in fname and "signed" in fname:
                return 1
    return 0

def filter_out_unsuccessful_jobs(jobs, client: ServiceTitanClient):
    unsuccessful_tags = [tag.get("id") for tag in fetch.get_tag_types(client) if "Unsuccessful" in tag.get("name")]
    return [job for job in jobs if unsuccessful_tags[0] not in job.get("tagTypeIds")]

# def process_completed_prefetches() -> None:
#     """Check prefetch futures and move completed downloads into the cache.

#     This function runs on every script execution.  It iterates over
#     ``st.session_state.prefetch_futures`` and for each future that has
#     finished, retrieves the images, stores them in
#     ``st.session_state.prefetched``, and removes the future from the
#     dictionary.  Because Streamlit reruns the script on most user
#     interactions, this effectively polls the futures when the user
#     interacts with the page.
#     """
#     done_ids: List[str] = []
#     for job_id, fut in st.session_state.prefetch_futures.items():
#         if fut.done():
#             try:
#                 images = fut.result()
#             except Exception:
#                 images = {}
#             st.session_state.prefetched[job_id] = images
#             done_ids.append(job_id)
#     for job_id in done_ids:
#         st.session_state.prefetch_futures.pop(job_id, None)

def fetch_jobs_button_call(tenant_filter, start_date, end_date, job_status_filter, filter_unsucessful, custom_job_id=None):
    with st.spinner("Retrieving jobs..."):
        tenant_filter = tenant_filter.split(" ")[0].lower()
        st.session_state.current_tenant = tenant_filter

        if tenant_filter not in st.session_state.clients:
            st.session_state.clients[tenant_filter] = get_client(tenant_filter)

        if "supabase" not in st.session_state.clients:
            st.session_state.clients["supabase"] = get_supabase()

        client = st.session_state.clients.get(tenant_filter)

        if tenant_filter not in st.session_state.employee_lists:
            st.session_state.employee_lists[tenant_filter] = get_all_employee_ids(client)

        if custom_job_id:
            jobs = fetch.fetch_jobs(start_date, end_date, client, custom_job_id)
        else:
            jobs = fetch.fetch_jobs(start_date, end_date, client, status_filters=job_status_filter)
            if filter_unsucessful:
                jobs = filter_out_unsuccessful_jobs(jobs, client)

        # TODO: Add invoice & payment logic
        invoice_ids = format.get_invoice_ids(jobs)
        invoices = fetch.fetch_invoices(invoice_ids, client)
        payments = fetch.fetch_payments(invoice_ids, client)

        invoices = {invoice['id']: format.format_invoice(invoice) for invoice in invoices}
        payments = format.format_payments(payments)
        # payments = flatten_list([format.format_invoice(payment) for payment in payments])

        jobs = format.combine_job_data(jobs, invoices, payments)

        st.session_state.jobs = jobs
        st.session_state.current_index = 0
        st.session_state.prefetched = {}
        # st.session_state.prefetch_futures = {}
        # Kick off prefetch for the first three jobs
        fetch.schedule_prefetches(client)
        # Trigger an immediate rerun to process any completed futures
        st.rerun()