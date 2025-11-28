from __future__ import annotations

import datetime as _dt
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json
import requests
import time
import asyncio

import streamlit as st
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

from servicetitan_api_client import ServiceTitanClient
from supabase import Client

import modules.formatting as format
import modules.google_store as gs
import modules.helpers as helpers
import modules.tasks as tasks

ATTACHMENT_DOWNLOADER_URL = 'https://attachment-downloader-901775793617.australia-southeast1.run.app/'
# ATTACHMENT_DOWNLOADER_URL = 'http://0.0.0.0:8000'

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
        job = format.add_appt_info(job, first_appts[job['id']], modifier='first')
        job = format.add_appt_info(job, last_appts[job['id']], modifier='last')
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
def fetch_attachment_bytes(attachment_id: int, _client: ServiceTitanClient) -> bytes:
    """Download an attachment and return its raw bytes.

    The ``attachment_id`` is passed to the URL builder to form
    ``forms/v2/tenant/{tenant}/jobs/attachment/{id}``, which should
    return the binary content.
    """
    url = _client.build_url("forms", "jobs/attachment", resource_id=attachment_id)
    return _client.get(url)

def download_attachment_to_gcs(attachment_id: int, _client: ServiceTitanClient, gcs_bucket: str, gcs_blob: str) -> str:
    """Downloads an attachment to gcs and return a signed url to access
    """
    data = fetch_attachment_bytes(attachment_id, _client)
    url = gs.upload_bytes_to_gcs_signed(data, gcs_bucket, gcs_blob)
    # print("signed_url:")
    # print(url)
    return url

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

    GCS_BUCKET = 'doc-check-attachments'

    attachments = fetch_job_attachments(job_id, client)
    grouped_meta = helpers.group_attachments_by_type(attachments)
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
            fut = pool.submit(download_attachment_to_gcs, att_id, client, GCS_BUCKET, f'{job_id}/{filename}')
            future_map[fut] = (category, filename, file_date, file_by)
        for fut in as_completed(future_map):
            category, filename, file_date, file_by = future_map[fut]
            try:
                signed_url = fut.result()
            except Exception as e:
                print(f"EXCEPTION: {e}")
                signed_url = None
            result[category].append((filename, client.from_utc_string(file_date), file_by, signed_url))
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

def get_tag_types(client: ServiceTitanClient):
    url = client.build_url('settings', 'tag-types')
    return client.get_all(url)

def request_job_download(job_id, tenant, base_url=ATTACHMENT_DOWNLOADER_URL, force_refresh=False):
    print(f'requested job download for {job_id}...')
    url = base_url + '/tasks/process-job'
    tasks.create_task(url, job_id, tenant, force_refresh)
    # payload = {
    #     "job_id": job_id,
    #     "tenant": tenant,
    #     "force_refresh": force_refresh
    # }

    # headers = {
    #     "accept": "application/json",
    #     "Content-Type": "application/json"
    # }

    # response = requests.post(url, json=payload, headers=headers)
    print(f'finished job download for {job_id}')
    return 

def schedule_prefetches(client: ServiceTitanClient, downloader_url=ATTACHMENT_DOWNLOADER_URL) -> None:
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
    end = min(current+5, len(jobs))
    for i in range(current, end):
        job_id = str(jobs[i].get("id"))
        executor = ThreadPoolExecutor(max_workers=5)
        future = executor.submit(request_job_download, job_id, st.session_state.current_tenant, downloader_url)

    return
        # Skip if already prefetched or scheduled
        # if job_id in st.session_state.prefetched or job_id in st.session_state.prefetch_futures:
        #     continue
        # # Schedule a background download
        # st.session_state.prefetch_futures[job_id] = future

# @st.cache_data(show_spinner=False)
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

# @st.cache_data(show_spinner=False)
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

    params = {
        'appliedToInvoiceIds': ','.join(invoice_ids)
    }

    payments = _client.get_all(base_path, params=params)
    return payments

def fetch_tag_types(client: ServiceTitanClient):
    url = client.build_url('settings', 'tag-types')
    return client.get_all(url)

def get_job_status(job_id: int, client: Client, tenant: str):
    """
    Return one of: {-1,0,1,2} representing 'error', 'pending', 'processing', 'processed', respectively or None if record doesn't exist.
    """
    response = (
        client.table("gcs_job_attachment_status")
        .select("status, error_msg, last_update")
        .eq("job_id", job_id)
        # .eq("tenant", tenant)
        .execute()
    )
    try:
        if len(response.data) == 0:
            return 0, "", None
        if len(response.data) > 1:
            return -1, response.data[0]['error_msg'], None
        else:
            return response.data[0]['status'], response.data[0].get('error_msg'), response.data[0].get('last_update')
    except KeyError:
        return None, None, None

def get_attachments_supabase(job_id: int, client: Client, tenant: str):
    response = (
        client.table("gcs_attachments")
        .select("job_id,type,url,file_date,file_by,file_name")
        .eq("job_id", int(job_id))
        # .eq("tenant", tenant)
        .execute()
    )
    return response.data

