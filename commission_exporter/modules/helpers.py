from __future__ import annotations

from datetime import date, timedelta, datetime, time
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json
from functools import reduce
import pandas as pd

import streamlit as st
import streamlit_authenticator as stauth
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

import holidays

from servicetitan_api_client import ServiceTitanClient
import modules.google_store as gs

import modules.data_formatting as format
import modules.data_fetching as fetching
import modules.lookup_tables as lookup
from pprint import pprint

def flatten_list(
        nested_list
    ):
    return [item for sublist in nested_list for item in sublist]

def get_client(
        tenant
    ) -> ServiceTitanClient:
    @st.cache_resource(show_spinner=False)
    def _create_client(tenant) -> ServiceTitanClient:
            client = ServiceTitanClient(
                app_key=gs.get_secret("st_app_key_tester"), 
                app_guid=gs.get_secret("st_servco_integrations_guid"), 
                tenant=gs.get_secret(f"st_tenant_id_{tenant}"), 
                client_id=gs.get_secret(f"st_client_id_{tenant}"), 
                client_secret=gs.get_secret(f"st_client_secret_{tenant}"), 
                environment="production"
            )
            return client
    return _create_client(tenant)

def get_all_employee_ids(
        client: ServiceTitanClient
    ):
    # roles_url = client.build_url("settings", "user-roles")
    # sales_codes = get_sales_codes(client.get_all(roles_url))
    # role_codes = format.get_sales_and_installer_codes(client.get_all(roles_url))
    tech_url = client.build_url("settings", "technicians")
    tech_params = {
        'active': 'Any'
    }
    tech_resp = client.get_all(tech_url, params=tech_params)
    # pprint(tech_resp)
    techs = format.format_employee_list(tech_resp)
    emp_url = client.build_url("settings", "employees")
    emp_params = {
        'active': 'Any'
    }
    office = format.format_employee_list(client.get_all(emp_url, params=emp_params))
    return techs | office

def get_sales_codes(
        roles_reponse
    ):
    sales_codes = set()
    for role in roles_reponse:
        if role['name'] == 'Technician - Sales':
            sales_codes.add(role['id'])
    return sales_codes

# def filter_out_unsuccessful_jobs(jobs, client: ServiceTitanClient):
#     unsuccessful_tags = [tag.get("id") for tag in fetching.fetch_tag_types(client) if "Unsuccessful" in tag.get("name")]
#     return [job for job in jobs if unsuccessful_tags[0] not in job.get("tagTypeIds")]

# def get_complaint_tags(jobs, client: ServiceTitanClient) -> List[int]:
#     tag_types = fetching.fetch_tag_types(client)
#     complaint_tags = [tag.get("id") for tag in tag_types if "complaint" in tag.get("name").lower()]
#     return complaint_tags

def check_payment_dates(
        job, 
        end_date
    ):
    if job.get("payment_dates"):
        if type(job.get("payment_dates")) != str:
            return True
        for date in job.get("payment_dates").split(', '):
            if datetime.strptime(date, "%Y-%m-%d").date() > end_date:
                return False
    return True

def categorise_job(
        job, 
        end_date,
        relevant_holidays
    ):

    def _categorise(
            job, 
            prefix
        ):
        status = job['status']
        balance = float(job['balance'])
        payments_in_time = job['payments_in_time']
        completed_dt = job['completed_dt']
        if job['unsuccessful']:
            return prefix + '_unsuccessful'
        if status == 'Completed':
            if completed_dt:
                if completed_dt.date() > end_date:
                    return prefix + '_wo'
            if balance <= 0 and payments_in_time:
                return prefix + '_complete_paid'
            if balance > 0 or not payments_in_time:
                return prefix + '_complete_unpaid'
        if status == 'Hold':
            return prefix + '_wo'
            # return 'wk_hold'
        if status == 'InProgress':
            return prefix + '_wo'
            # return 'wk_progress'
        if status == 'Scheduled':
            return prefix + '_wo'
            # return 'wk_scheduled'
        return prefix + '_uncategorised'
    
    day = job['first_appt_start_dt'].weekday()
    if day <5: # weekdays
        if job['first_appt_start_dt'].date() in relevant_holidays:
            return _categorise(job, 'ph')
        if job['first_appt_start_dt'].time() >= time(18,0,0) and job['first_appt_start_dt'].date() == job.get('completed_dt', datetime(2000,1,1)).date():
            return _categorise(job, 'ah')
        return _categorise(job, 'wk')
    if day >=5: # weekends
        return _categorise(job, 'wkend')

def merge_dfs(
        dfs: list, 
        on='job_id', 
        how='left'
    ):
    return reduce(lambda left, right: pd.merge(left, right, on=on, how=how), dfs)

def get_last_day_of_month_datetime(
        year, 
        month
    ):
    """
    Returns the last day of the specified month and year as a date object.
    """
    if month == 12:
        next_month_year = year + 1
        next_month = 1
    else:
        next_month_year = year
        next_month = month + 1
    
    first_day_of_next_month = date(next_month_year, next_month, 1)
    last_day_of_current_month = first_day_of_next_month - timedelta(days=1)
    return last_day_of_current_month

def get_dates_in_month_datetime(
        year, 
        month
    ) -> List[date]:
    start_date = date(year, month, 1)
    dates_in_month = []
    current_date = start_date
    while current_date.month == month:
        dates_in_month.append(current_date)
        current_date += timedelta(days=1)
    return dates_in_month

def get_public_holidays(
        state: str
    ):
    national = holidays.Australia()
    state_hols = holidays.Australia(subdiv=state)

    combined = holidays.HolidayBase()
    combined.update(national)
    combined.update(state_hols)
    return combined

def get_threshold_days(
        dates: List[date], 
        holidays: List[date]
    ):
    threshold_days = 0
    for day in dates:
        if day not in holidays and day.weekday() not in [5,6]:
            threshold_days += 1
    return threshold_days

def get_holidays(state: str):
    return holidays.Australia(state=state)

def australian_public_holidays_between(
        start_date: date,
        end_date: date,
        state: str
    ) -> List[date]:
    """
    Returns a list of dates for Australian public holidays
    between start_date and end_date inclusive, for a given state.

    Parameters
    ----------
    start_date : datetime.date
    end_date   : datetime.date
    state      : str
        Australian state/territory code:
        'NSW', 'VIC', 'QLD', 'SA', 'WA', 'TAS', 'ACT', 'NT'

    Returns
    -------
    List[datetime.date]
    """

    if start_date > end_date:
        raise ValueError("start_date must be before or equal to end_date")

    au_holidays = holidays.Australia(state=state)

    results = []
    current = start_date
    while current <= end_date:
        if current in au_holidays:
            results.append(current)
        current += timedelta(days=1)
    return results