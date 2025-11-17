"""
Streamlit application for browsing ServiceTitan jobs within a date range.

This app allows a user to specify a start and end date, retrieves jobs
created within that window, and then displays the images attached to
each job.  To improve performance, the app prefetches images for the
next two jobs while the current job is being viewed so that
navigation feels instant.  Each job also has a simple form in the
sidebar with seven checkboxes; when submitted, a PATCH request is
issued to a URL of your choosing with the checkbox values.

The attachment API endpoints used here assume the following
structure:

```
forms/v2/tenant/{tenant}/jobs/{job_id}/attachments           # to list attachments
forms/v2/tenant/{tenant}/jobs/attachment/{attachment_id}     # to download an attachment
jpm/v2/tenant/{tenant}/jobs                                  # to list jobs
```

If your ServiceTitan account exposes different endpoints or requires
additional query parameters to filter by date, adjust the
``fetch_jobs`` function accordingly.  The ServiceTitan developer
documentation mentions that you can “find all jobs that have work
scheduled within a given date range” and “find all jobs completed
within a date range”【503339604690427†L129-L134】; however, the exact query
parameters depend on your account configuration.

References
----------
* The ServiceTitan developer docs recommend authenticating via
  OAuth2 client credentials and including the application key in the
  ``ST-App-Key`` header【861456388967273†L72-L119】【861456388967273†L122-L140】.
* The Job Planning overview notes that the API can list jobs
  scheduled or completed within a date range【503339604690427†L129-L134】.
"""

from __future__ import annotations

import datetime as _dt
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json
from google.cloud import secretmanager

import streamlit as st
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

from servicetitan_api_client import ServiceTitanClient


###############################################################################
# Configuration and helpers
###############################################################################

# File extensions considered images.  Extend as needed.
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

def get_client(state) -> ServiceTitanClient:
    @st.cache_resource
    def _create_client(state) -> ServiceTitanClient:
            state_code = state_codes()[state]
            client = ServiceTitanClient(
                app_key=get_secret("ST_app_key_tester"), 
                tenant=get_secret(f"ST_tenant_id_{state_code}"), 
                client_id=get_secret(f"ST_client_id_{state_code}"), 
                client_secret=get_secret(f"ST_client_secret_{state_code}"), 
                environment="production"
            )
            return client
    return _create_client(state)


@st.cache_data(show_spinner=False)
def fetch_jobs(
    start_date: _dt.date,
    end_date: _dt.date,
    _client: ServiceTitanClient,
    job_id: str = None,
) -> List[Dict[str, Any]]:
    """
    Retrieve all jobs created between `start_date` and `end_date`,
    converting the local date boundaries into UTC timestamps. If
    job_num specified, just fetches that job.
    """
    tenant = _client.tenant or "{tenant}"
    base_path = f"jpm/v2/tenant/{tenant}/jobs"
    page = 1
    jobs: List[Dict[str, Any]] = []

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
        for job in page_data:
            created = job.get("createdOn")
            if not created:
                return []
            try:
                dt = _dt.datetime.fromisoformat(created.replace("Z", "+00:00")).date()
            except Exception:
                return []
            jobs.append(job)
        return jobs

    while True:
        # Convert local dates to UTC ISO strings
        created_after = _client.start_of_day_utc_string(start_date)
        created_before = _client.end_of_day_utc_string(end_date)

        params = {
            "page": page,
            "pageSize": 50,
            "createdOnOrAfter": created_after,
            "createdBefore": created_before,
        }
        try:
            resp = _client.get(base_path, params=params)
        except Exception:
            break
        if not isinstance(resp, dict):
            break
        page_data: Iterable[Dict[str, Any]] = resp.get("data") or []
        for job in page_data:
            created = job.get("createdOn")
            if not created:
                continue
            try:
                dt = _dt.datetime.fromisoformat(created.replace("Z", "+00:00")).date()
            except Exception:
                continue
            if start_date <= dt <= end_date:
                jobs.append(job)
        # Pagination logic (unchanged)
        has_more = resp.get("hasMore")
        if has_more:
            page += 1
            continue
        break
    return jobs


@st.cache_data(show_spinner=False)
def fetch_job_attachments(job_id: str, _client: ServiceTitanClient) -> List[Dict[str, Any]]:
    """Retrieve attachment metadata for the given job ID.

    This function calls the attachments listing endpoint for a job and
    returns a list of attachment objects (dictionaries).  Each
    attachment is expected to contain an ``id`` and ``fileName`` key.
    """
    attachments_url = _client.build_url("forms", f"jobs/{job_id}/attachments", version=2)
    data = _client.get(attachments_url)
    attachments = data.get("data", []) if isinstance(data, dict) else []
    return attachments


@st.cache_data(show_spinner=False)
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
        file_name = att.get("fileName") or att.get("filename") or att.get("name")
        att_id = att.get("id")
        if not file_name or att_id is None:
            continue

        # Extract the lower‑cased extension with a leading dot
        ext = file_name.lower().rsplit(".", 1)[-1] if "." in file_name else ""
        ext_with_dot = f".{ext}" if ext else ""
        for category, exts in extension_map.items():
            if ext_with_dot in exts:
                result[category].append((file_name, int(att_id)))
                break  # stop after the first matching category
    return result


def download_images_for_job(job_id: str, client: ServiceTitanClient) -> List[Tuple[str, bytes]]:
    """Download all image attachments for a job and return them.

    This helper performs two API calls: one to list attachments and a
    second to download each image concurrently.  It does not call
    Streamlit functions and therefore may be safely executed in a
    background thread.
    """
    attachments = fetch_job_attachments(job_id, client)
    meta = filter_image_attachments(attachments)
    images: List[Tuple[str, bytes]] = []
    if not meta:
        return images
    # Use a local ThreadPoolExecutor to download images in parallel
    max_workers = min(8, len(meta)) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        future_to_name: Dict[Future[bytes], str] = {}
        for filename, att_id in meta:
            fut = pool.submit(fetch_image_bytes, att_id, client)
            future_to_name[fut] = filename
        for fut in as_completed(future_to_name):
            filename = future_to_name[fut]
            try:
                data = fut.result()
            except Exception:
                continue
            images.append((filename, data))
    return images

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

###############################################################################
# Streamlit app logic
###############################################################################

def main() -> None:
    st.set_page_config(page_title="ServiceTitan Job Browser", layout="wide")
    st.title("ServiceTitan Job Image Browser")
    client = get_client('NSW')

    doc_checks = get_doc_check_criteria()

    # Initialise session state collections
    if "jobs" not in st.session_state:
        st.session_state.jobs: List[Dict[str, Any]] = []
    if "current_index" not in st.session_state:
        st.session_state.current_index: int = 0
    if "prefetched" not in st.session_state:
        st.session_state.prefetched: Dict[str, List[Tuple[str, bytes]]] = {}
    if "prefetch_futures" not in st.session_state:
        st.session_state.prefetch_futures: Dict[str, Future] = {}
    if "app_guid" not in st.session_state:
        st.session_state.app_guid = get_secret('ST_servco_integrations_guid')

    # Sidebar controls for date range and patch URL
    with st.sidebar:
        st.header("Job date range")
        today = _dt.date.today()
        default_start = today - _dt.timedelta(days=7)
        start_date = st.date_input("Start date", value=default_start)
        end_date = st.date_input("End date", value=today)
        custom_job_id = st.text_input(
            "Job ID Search", placeholder="Manual search for job", help="Job ID is different to the job number. ID is the number at the end of the URL of the job's page in ServiceTitan"
        )
        fetch_jobs_button = st.button("Fetch Jobs", type="primary")
        st.markdown("---")
        st.header("Form submission")

    # When the fetch button is pressed, call the API and reset state
    if fetch_jobs_button:
        with st.spinner("Retrieving jobs..."):
            if custom_job_id:
                jobs = fetch_jobs(start_date, end_date, client, custom_job_id)
            else:
                jobs = fetch_jobs(start_date, end_date, client)
        st.session_state.jobs = jobs
        st.session_state.current_index = 0
        st.session_state.prefetched = {}
        st.session_state.prefetch_futures = {}
        # Kick off prefetch for the first three jobs
        _schedule_prefetches(client)
        # Trigger an immediate rerun to process any completed futures
        st.rerun()

    # Process completed prefetch futures and update prefetched cache
    _process_completed_prefetches()

    # Display the current job if available
    if st.session_state.jobs:
        idx = st.session_state.current_index
        if idx < 0:
            idx = 0
        if idx >= len(st.session_state.jobs):
            idx = len(st.session_state.jobs) - 1
        job = st.session_state.jobs[idx]
        job_id = str(job.get("id"))
        job_num = str(job.get("jobNumber"))


        # Sidebar form for the current job
        with st.sidebar.form(key=f"form_{job_num}"):
            st.subheader(f"Job {job_num} form")
            checks = {}
            # initial_bits = get_job_external_data(job_id, client, st.session_state.app_guid)
            initial_checks = get_job_external_data(job_id, client, st.session_state.app_guid)
            for check_code, check in doc_checks.items():
                default = bool(initial_checks.get(check_code, False))
                checks[check_code] = int(st.checkbox(check, key=f"{job_num}_{check_code}", value=default))

            # for i in range(1, 8): # TODO: Might want to change this logic to make it more robust to future changes
            #     default = initial_bits and len(initial_bits) >= i and initial_bits[i-1] == "1"
            #     checks[f"check{i}"] = st.checkbox(f"Check {i}", key=f"{job_num}_check{i}", value=default)
            # for i in range(1, 8):
            #     checks[f"check{i}"] = st.checkbox(f"Check {i}", key=f"{job_num}_check{i}")
            submitted = st.form_submit_button("Submit")
            if submitted:
                # Prepare payload and send PATCH request
                # encoded = ''.join(['1' if checks[f"check{i}"] else '0' for i in range(1, 8)])
                encoded = json.dumps(checks)
                print(encoded)
                external_data_payload = {
                    "externalData": {
                        "patchMode": "Replace",
                        "applicationGuid": st.session_state.app_guid,
                        "externalData": [{"key": "docchecks", "value": encoded}]
                    }
                }

                patch_url = client.build_url('jpm', 'jobs', resource_id=job_id)
                try:
                    client.patch(patch_url, json=external_data_payload)
                    st.success("Form submitted successfully")
                except Exception as e:
                    st.error(f"Failed to submit form: {e}")

        # Navigation buttons
        col_prev, col_next = st.columns([1, 1])
        with col_prev:
            if st.button("Previous Job"):
                if st.session_state.current_index > 0:
                    st.session_state.current_index -= 1
                    _schedule_prefetches(client)
                    st.rerun()
        with col_next:
            if st.button("Next Job"):
                if st.session_state.current_index < len(st.session_state.jobs) - 1:
                    st.session_state.current_index += 1
                    _schedule_prefetches(client)
                    st.rerun()

        # Display job details and images
        st.write(f"**Viewing job {job_num} ({idx + 1} of {len(st.session_state.jobs)})**")
        # Show some basic job info if available
        st.json(job, expanded=False)

        images = st.session_state.prefetched.get(job_id)
        if images is None:
            # If not already prefetched, download synchronously
            with st.spinner("Downloading images..."):
                images = download_images_for_job(job_id, client)
            st.session_state.prefetched[job_id] = images

        if images:
            cols = st.columns(3)
            for idx_img, (filename, data) in enumerate(images):
                with cols[idx_img % 3]:
                    st.image(data, caption=filename, width=150)
        else:
            st.info("No image attachments for this job.")

    else:
        st.info("Enter a date range and click 'Fetch Jobs' to begin.")


def _schedule_prefetches(client: ServiceTitanClient) -> None:
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
        future = executor.submit(download_images_for_job, job_id, client)
        st.session_state.prefetch_futures[job_id] = future


def _process_completed_prefetches() -> None:
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


if __name__ == "__main__":
    main()