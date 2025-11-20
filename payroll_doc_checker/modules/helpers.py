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

def add_appt_info(job, appt, modifier='first'):
    job[f'{modifier}_appt_start'] = appt.get("start")
    job[f'{modifier}_appt_end'] = appt.get("end")
    job[f'{modifier}_appt_arrival_start'] = appt.get("arrivalWindowStart")
    job[f'{modifier}_appt_arrival_end'] = appt.get("arrivalWindowEnd")
    job[f'{modifier}_appt_num'] = appt.get("appointmentNumber")
    return job

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
        if _client.app_guid:
            params["externalDataApplicationGuid"] = _client.app_guid
        try:
            resp = _client.get(base_path, params=params)
        except Exception:
            return []
        if not isinstance(resp, dict):
            return []
        jobs: Iterable[Dict[str, Any]] = resp.get("data") or []
    else:
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

    first_appt_ids = [str(job.get("firstAppointmentId")) for job in jobs]
    last_appt_ids = [str(job.get("lastAppointmentId")) for job in jobs]
    appt_url = _client.build_url('jpm', 'appointments')
    first_appts = _client.get_all_id_filter(appt_url, first_appt_ids)
    first_appts = {appt.get("jobId"): appt for appt in first_appts}
    last_appts = _client.get_all_id_filter(appt_url, last_appt_ids)
    last_appts = {appt.get("jobId"): appt for appt in last_appts}
    for job in jobs:
        job = add_appt_info(job, first_appts[job['id']], modifier='first')
        job = add_appt_info(job, last_appts[job['id']], modifier='last')
    return jobs


# @st.cache_data(show_spinner=False)
def fetch_job_attachments(job_id: str, _client: ServiceTitanClient) -> List[Dict[str, Any]]:
    """Retrieve attachment metadata for the given job ID.

    This function calls the attachments listing endpoint for a job and
    returns a list of attachment objects (dictionaries).  Each
    attachment is expected to contain an ``id`` and ``fileName`` key.
    """
    attachments_url = _client.build_url("forms", f"jobs/{job_id}/attachments", version=2)
    attachments = _client.get_all(attachments_url)
    # attachments = data.get("data", []) if isinstance(data, dict) else []
    return attachments


# @st.cache_data(show_spinner=False)
def fetch_image_bytes(attachment_id: int, _client: ServiceTitanClient) -> bytes:
    """Download an attachment and return its raw bytes.

    The ``attachment_id`` is passed to the URL builder to form
    ``forms/v2/tenant/{tenant}/jobs/attachment/{id}``, which should
    return the binary content.
    """
    url = _client.build_url("forms", "jobs/attachment", resource_id=attachment_id)
    return _client.get(url)


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


def download_attachments_for_job(job_id: str, client: ServiceTitanClient) -> Dict[str, List[Tuple[str, Any]]]:
    """Download all attachments for a job and group them by type.

    This helper performs two API calls: one to list attachments and a
    second to download each attachment.  It groups attachments into
    ``imgs`` and ``pdfs`` by default (using the same extension map as
    :func:`group_attachments_by_type`).  Each entry in the returned
    dictionary is a list of ``(filename, data)`` tuples, where ``data``
    is the raw bytes of the attachment.  If no attachments exist for a
    category, the list will be empty.
    """
    attachments = fetch_job_attachments(job_id, client)
    grouped_meta = group_attachments_by_type(attachments)
    result: Dict[str, List[Tuple[str, Any]]] = {key: [] for key in grouped_meta}
    # Download images and PDFs.  Use a ThreadPoolExecutor to parallelise
    # downloads across all attachments regardless of type.
    tasks: List[Tuple[str, int, str]] = []  # (category, att_id, filename)
    for category, items in grouped_meta.items():
        for filename, att_id, file_date, file_by in items:
            tasks.append((category, att_id, filename, file_date, file_by))
    if not tasks:
        return result
    max_workers = min(8, len(tasks)) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        future_map: Dict[Future[bytes], Tuple[str, str]] = {}
        for category, att_id, filename, file_date, file_by in tasks:
            fut = pool.submit(fetch_image_bytes, att_id, client)
            future_map[fut] = (category, filename, file_date, file_by)
        for fut in as_completed(future_map):
            category, filename, file_date, file_by = future_map[fut]
            try:
                data = fut.result()
            except Exception:
                data = None
            result[category].append((filename, client.from_utc_string(file_date), file_by, data))
    return result

def get_job_external_data(job, key="docchecks_testing"):
    external_entries = job.get("externalData", [])
    for entry in external_entries:
        if entry.get("key") == key:
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

def pre_fill_quote_signed_check(pdfs):
    for pdf in pdfs:
        fname = pdf[0].lower()
        if "estimate" in fname and "signed" in fname:
            return 1
    return 0

def pre_fill_invoice_signed_check(pdfs):
    for pdf in pdfs:
        fname = pdf[0].lower()
        if "invoice" in fname and "signed" in fname:
            return 1
    return 0

def get_tag_types(client: ServiceTitanClient):
    url = client.build_url('settings', 'tag-types')
    return client.get_all(url)

def filter_out_unsuccessful_jobs(jobs, client: ServiceTitanClient):
    unsuccessful_tags = [tag.get("id") for tag in get_tag_types(client) if "Unsuccessful" in tag.get("name")]
    return [job for job in jobs if unsuccessful_tags[0] not in job.get("tagTypeIds")]

def schedule_prefetches(client: ServiceTitanClient) -> None:
    """Ensure up to three jobs (current and next two) are prefetched.

    For the current job index ``i``, this function schedules
    prefetches for jobs at ``i``, ``i+1``, and ``i+2``.  Already
    completed downloads are not rescheduled.  Existing futures are
    left to finish.
    """
    jobs = st.session_state.jobs
    if not jobs:
        return
    current = st.session_state.current_index
    end = min(current + 3, len(jobs))
    for i in range(current, end):
        job_id = str(jobs[i].get("id"))
        # Skip if already prefetched or scheduled
        if job_id in st.session_state.prefetched or job_id in st.session_state.prefetch_futures:
            continue
        # Schedule a background download
        executor = ThreadPoolExecutor(max_workers=1)
        future = executor.submit(download_attachments_for_job, job_id, client)
        st.session_state.prefetch_futures[job_id] = future


def process_completed_prefetches() -> None:
    """Check prefetch futures and move completed downloads into the cache.

    This function runs on every script execution.  It iterates over
    ``st.session_state.prefetch_futures`` and for each future that has
    finished, retrieves the images, stores them in
    ``st.session_state.prefetched``, and removes the future from the
    dictionary.  Because Streamlit reruns the script on most user
    interactions, this effectively polls the futures when the user
    interacts with the page.
    """
    done_ids: List[str] = []
    for job_id, fut in st.session_state.prefetch_futures.items():
        if fut.done():
            try:
                images = fut.result()
            except Exception:
                images = []
            st.session_state.prefetched[job_id] = images
            done_ids.append(job_id)
    for job_id in done_ids:
        st.session_state.prefetch_futures.pop(job_id, None)

def fetch_jobs_button_call(tenant_filter, start_date, end_date, job_status_filter, filter_unsucessful, custom_job_id=None):
    with st.spinner("Retrieving jobs..."):
        tenant_filter = tenant_filter.split(" ")[0].lower()
        st.session_state.current_tenant = tenant_filter

        if tenant_filter not in st.session_state.clients:
            st.session_state.clients[tenant_filter] = get_client(tenant_filter)

        client = st.session_state.clients.get(tenant_filter)

        if tenant_filter not in st.session_state.employee_lists:
            st.session_state.employee_lists[tenant_filter] = get_all_employee_ids(client)

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
        # Kick off prefetch for the first three jobs
        schedule_prefetches(client)
        # Trigger an immediate rerun to process any completed futures
        # st.rerun()