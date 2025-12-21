from __future__ import annotations

import streamlit as st
import datetime as _dt
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable

from servicetitan_api_client import ServiceTitanClient

# @st.cache_data(show_spinner=False)
def fetch_jobs(
    _client: ServiceTitanClient,
    start_date: _dt.date = None,
    end_date: _dt.date = None,
    job_id_ls: list = None,
) -> List[Dict[str, Any]]:
    """
    Retrieve all jobs created between `start_date` and `end_date`,
    converting the local date boundaries into UTC timestamps.
    """

    if not start_date and not end_date and not job_id_ls:
        return []

    base_path = _client.build_url('jpm', 'jobs')
    jobs: List[Dict[str, Any]] = []
    params = {}
    if _client.app_guid:
        params["externalDataApplicationGuid"] = _client.app_guid

    if start_date and end_date:
        created_after = _client.start_of_day_utc_string(start_date)
        created_before = _client.end_of_day_utc_string(end_date)

        params["createdOnOrAfter"] = created_after
        params["createdBefore"] = created_before

        return _client.get_all(base_path, params=params)

    if job_id_ls:
        return _client.get_all_id_filter(base_path, job_id_ls, params=params)
    
    return []

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
        # 'appliedToInvoiceIds': ','.join(invoice_ids)
    }

    payments = _client.get_all_id_filter(base_path, ids= invoice_ids, id_filter_name='appliedToInvoiceIds', params=params)
    return payments

def fetch_tag_types(client: ServiceTitanClient):
    url = client.build_url('settings', 'tag-types')
    return client.get_all(url)

# @st.cache_data(show_spinner=False)
def fetch_appt_assmnts(
    _client: ServiceTitanClient,
    start_date: _dt.date = None,
    end_date: _dt.date = None,
    job_id: str | None = None,
    appt_ids: List[str] = None
) -> List[Dict[str, Any]]:
    """
    Retrieve all appointment assignments created between `start_date` and `end_date`,
    converting the local date boundaries into UTC timestamps.
    """

    base_path = _client.build_url('dispatch', 'appointment-assignments')

    if start_date and end_date:
        created_after = _client.start_of_day_utc_string(start_date)
        created_before = _client.end_of_day_utc_string(end_date)
        params = {
                    "createdOnOrAfter": created_after,
                    "createdBefore": created_before,
                }
        return _client.get_all(base_path, params=params)
    
    # If job_id specified, only return that job
    if job_id:
        params = {
            "jobId": job_id,
        }
        try:
            resp = _client.get(base_path, params=params)
        except Exception:
            return []
        if not isinstance(resp, dict):
            return []
        page_data: Iterable[Dict[str, Any]] = resp.get("data") or []
        return page_data
    
    if appt_ids:
        params = {
            'appointmentIds': ','.join[appt_ids]
        }
        return _client.get_all(base_path, params=params)
    
    return []

# @st.cache_data(show_spinner=False)
def fetch_appts(
    _client: ServiceTitanClient,
    start_date: _dt.date,
    end_date: _dt.date,
    # job_id: str | None = None
) -> List[Dict[str, Any]]:
    """
    Retrieve all appointments created between `start_date` and `end_date`,
    converting the local date boundaries into UTC timestamps.
    """

    base_path = _client.build_url('jpm', 'appointments')

    starts_after = _client.start_of_day_utc_string(start_date)
    starts_before = _client.end_of_day_utc_string(end_date)
    # If job_id specified, only return that job
    # if job_id:
    #     params = {
    #         "jobId": job_id,
    #     }
    #     try:
    #         resp = _client.get(base_path, params=params)
    #     except Exception:
    #         return []
    #     if not isinstance(resp, dict):
    #         return []
    #     page_data: Iterable[Dict[str, Any]] = resp.get("data") or []
    #     return page_data
    params = {
                "startsOnOrAfter": starts_after,
                "startsBefore": starts_before,
            }
    appts = _client.get_all(base_path, params=params)
    return appts

# @st.cache_data(show_spinner=False)
def fetch_estimates(
    start_date: _dt.date,
    end_date: _dt.date,
    _client: ServiceTitanClient,
    job_id: str | None = None
) -> List[Dict[str, Any]]:
    """
    Retrieve all estimates created between `start_date` and `end_date`,
    converting the local date boundaries into UTC timestamps.
    """

    base_path = _client.build_url('sales', 'estimates')

    created_after = _client.start_of_day_utc_string(start_date)
    created_before = _client.end_of_day_utc_string(end_date)
    # If job_id specified, only return that job
    if job_id:
        params = {
            "jobId": job_id,
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
    ests = _client.get_all(base_path, params=params)
    return ests