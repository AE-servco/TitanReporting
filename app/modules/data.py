import pandas as pd
from google.cloud import secretmanager
from datetime import datetime, time, date, timedelta
from zoneinfo import ZoneInfo
from io import BytesIO
from PIL import Image
import pytz
import holidays
from concurrent.futures import ThreadPoolExecutor
import threading
from streamlit import session_state as ss

import servicepytan as sp

def flatten_list(nested_list):
    return [item for sublist in nested_list for item in sublist]

def clear_ss():
    for key in ss.keys():
        ss[key] = None

def format_employee_list(employee_response):
    # input can be either technician response or employee response
    formatted = {}
    for employee in employee_response:
        formatted[employee['id']] = employee['name']
    return formatted

def get_all_employee_ids(state):
    check_and_update_ss_for_data_service(state)
    techs = format_employee_list(ss[f'st_data_service_{state}'].get_all_technicians())
    office = format_employee_list(ss[f'st_data_service_{state}'].get_api_data('settings', 'employees', options={'active': 'Any'}))
    return techs | office


def format_ST_date(ST_date, format_str='%Y-%m-%dT%H:%M:%S'):
    return sp.convert_ST_datetime_to_local_str(ST_date, local_tz="Australia/Sydney", format_str=format_str)

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

def get_public_holidays(state):
    national = holidays.Australia()
    state_hols = holidays.Australia(subdiv=state)

    combined = holidays.HolidayBase()
    combined.update(national)
    combined.update(state_hols)
    return combined

def check_dates_for_hols(date_range, holidays):
    holidays_in_range = set()
    date_iter = date_range[0]
    date_end = date_range[1] + timedelta(days=1)
    while date_iter != date_end:
        if date_iter in holidays:
            holidays_in_range.add(date_iter)
        date_iter += timedelta(days=1)
    return holidays_in_range

def get_data_service(state):
    state_code = state_codes()[state]
    st_conn = sp.auth.servicepytan_connect(app_key=get_secret("ST_app_key_tester"), tenant_id=get_secret(f"ST_tenant_id_{state_code}"), client_id=get_secret(f"ST_client_id_{state_code}"), 
    client_secret=get_secret(f"ST_client_secret_{state_code}"), timezone="Australia/Sydney")
    st_data_service = sp.DataService(conn=st_conn)

    return st_data_service

def check_and_update_ss_for_data_service(state):
    if f'st_data_service_{state}' not in ss or f'st_data_service_{state}' is None:
        ss[f'st_data_service_{state}'] = get_data_service(state)

def get_invoices_for_xero(state, start_date, end_date):

    def account_codes():
        codes = {
            'NSW': '210-2',
            'WA': '210-1',
        }
        return codes

    def format_address(customerAddress):
        address = f"{customerAddress['street']}, {customerAddress['city']}, {customerAddress['state']} {customerAddress['zip']}, {customerAddress['country']}"
        if customerAddress['unit']:
            address = f"{customerAddress['unit']}/{address}"
        return address

    def format_invoice(invoice):
        formatted = {}
        invoice_num = invoice['referenceNumber']
        if state == "WA" and invoice_num.startswith("1"):
            formatted['*InvoiceNumber'] = 'W' + invoice_num
        else:
            formatted['*InvoiceNumber'] = invoice_num
        formatted['Invoice Date'] = datetime.fromisoformat(invoice['invoiceDate'].replace('Z', '+00:00')).strftime("%m/%d/%Y")
        formatted['*ContactName'] = invoice['customer']['name']
        formatted['Location Address'] = format_address(invoice['customerAddress'])
        formatted['*Description'] = invoice['job']['type']
        formatted['*UnitAmount'] = invoice['subTotal']
        formatted['*TaxType'] = "GST on Income"
        formatted['Sum'] = invoice['total']
        formatted['*AccountCode'] = account_codes()[state]
        return formatted

    check_and_update_ss_for_data_service(state)

    start_time = datetime.combine(start_date, time(0,0,0))
    end_time = datetime.combine(end_date, time(23,59,59))

    invoice_response = ss[f'st_data_service_{state}'].get_invoices_between(start_time, end_time)

    invoices = [format_invoice(invoice) for invoice in invoice_response]
    
    return pd.DataFrame(invoices)

def convert_df_for_download(df):
    if df is None:
        df = pd.DataFrame()
    return df.to_csv(index=False).encode("utf-8")


def get_commission_data(state, start_date, end_date):

    app_guid = get_secret('ST_servco_integrations_guid')
    unsucessfultag = 116255355
    # Filter for unsuccessful tag on job, soldBy should be in job, get technician name through settings endpoint
    # Payment type???
    # Do they get comms if its just the call out fee?

    def format_job(job, technicians):
        if 116255355 in job['tagTypeIds'] or job['jobStatus'] == 'Canceled': # Unsuccessful or cancelled 
            return None
        formatted = {}
        if job['soldById'] is not None:
            formatted['Sold By'] = technicians[job['soldById']]
        else:
            appts = ss[f'st_data_service_{state}'].get_appointment_assignments_by_job_id(job['id'])
            formatted['Sold By'] = ','.join([appt['technicianName'] for appt in appts]) + ' (Primary Tech)'
        # formatted['Sold By'] = technicians[job['soldById']] if job['soldById'] is not None else "None"
        # formatted['Primary Technician'] = invoice['customer']['name']
        # formatted['Created Date'] = job['createdOn'] if job['createdOn'] is not None else "None"
        formatted['Created Date'] = sp.convert_utc_datetime_to_local(sp.convert_ST_datetime_to_object(job['createdOn']), "Australia/Sydney").strftime("%m/%d/%Y")
        formatted['Completion Date'] = sp.convert_utc_datetime_to_local(sp.convert_ST_datetime_to_object(job['completedOn']), "Australia/Sydney").strftime("%m/%d/%Y") if job['completedOn'] is not None else "None"
        formatted['Job #'] = job['jobNumber'] if job['jobNumber'] is not None else "None"
        # formatted['Suburb'] = job['subTotal']
        # formatted['Jobs Subtotal'] = "GST on Income"
        # formatted['Payment Types'] = job['total']
        formatted['Status'] = job['jobStatus'] if job['jobStatus'] is not None else "None"
        formatted['invoiceId'] = job['invoiceId'] if job['invoiceId'] is not None else "None"
        formatted['externalData'] = job['externalData'][0] if job['externalData'] is not None and job['externalData'] != [] else "None"
        return formatted

    def format_invoice(invoice):
        formatted = {}
        formatted['Suburb'] = invoice['customerAddress']['city']
        formatted['Jobs Subtotal'] = invoice['subTotal']
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

    def get_invoice_ids(job_response):
        return [str(job['invoiceId']) for job in job_response]
    
    # def get_job_ids(job_response):
    #     return [str(job['id']) for job in job_response]

    check_and_update_ss_for_data_service(state)

    start_time = datetime.combine(start_date, time(0,0,0))
    end_time = datetime.combine(end_date, time(23,59,59))
 
    technicians_response = ss[f'st_data_service_{state}'].get_all_technicians()
    technicians = format_employee_list(technicians_response)

    job_response = ss[f'st_data_service_{state}'].get_jobs_created_between(start_time, end_time, app_guid=app_guid)
    jobs_w_nones = [format_job(job, technicians) for job in job_response]
    jobs = [job for job in jobs_w_nones if job is not None]
    jobs_df = pd.DataFrame(jobs)
    
    invoice_ids = get_invoice_ids(job_response)

    invoice_response = ss[f'st_data_service_{state}'].get_invoices_by_id(invoice_ids)
    invoices = [format_invoice(invoice) for invoice in invoice_response]
    invoices_df = pd.DataFrame(invoices)

    payments_response = ss[f'st_data_service_{state}'].get_payments_for_invoices(invoice_ids)
    payments = [format_payment(payment) for payment in payments_response]
    payments_flat = flatten_list(payments)
    payments_df = pd.DataFrame(payments_flat)
    payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(','.join)

    merged = pd.merge(pd.merge(jobs_df, invoices_df, on='invoiceId', how='left'), payments_grouped, on='invoiceId', how='left')

    merged_reordered = merged.loc[:, ['Sold By', 'Created Date', 'Job #', 'Suburb', 'Jobs Subtotal', 'Payment Types', 'Status', 'Completion Date', 'Payments', 'externalData']]

    return merged_reordered.sort_values(by=['Sold By', 'Status'])

# def get_full_commission_data(state, start_date, end_date):
#     # 25k threshold p week or 5k per day
#     # net sold, not gross sold - need materials and things - problem for later
#     # doc check - assume we have this per job from other app in the works
#     # OT/weekend/PH rates

#     # Order of Ops:
#     # 1. get weekly sold per plumber (put code in place that will be able to work out reductions)
#     #   - get all jobs created btwn dates
#     #   - get invoices attached to them
#     # 2. doc check checker (just put code in place, waiting on data from other report for full functionality)
#     # 3. check for payments
#     # 4. output similar to current spreadsheet so we can compare
#     #   - per plumber, per job, per status (completed, etc.)

#     public_hols = get_public_holidays(state)
#     public_hols_in_week = check_dates_for_hols((start_date, end_date), public_hols)

#     if public_hols_in_week:
#         threshold = 5000 * len(public_hols_in_week)
#     else:
#         threshold = 25000

#     def get_unsuccessful_tag():
#         tags = ss[f'st_data_service_{state}'].get_all_tag_types()
#         for tag in tags:
#             if 'Unsuccessful' in tag['name']:
#                 return tag['id']

#     def get_job_cost(invoice):
#         # TODO: will be used to get job cost when that data is available.
#         return 0.0

#     def format_job(job, technicians, unsuccessful_tag):
#         formatted = {}
#         if unsuccessful_tag in job['tagTypeIds']:
#             formatted['Status'] = "Unsuccessful"
#         else:
#             # formatted['unsuccessful'] = 0
#             formatted['Status'] = job['jobStatus'] if job['jobStatus'] is not None else "None"
#         if job['soldById'] is not None:
#             formatted['Sold By'] = technicians[job['soldById']]
#         else:
#             formatted['Sold By'] = None
#         formatted['Created Date'] = sp.convert_utc_datetime_to_local(sp.convert_ST_datetime_to_object(job['createdOn']), "Australia/Sydney").strftime("%m/%d/%Y")
#         formatted['Completion Date'] = sp.convert_utc_datetime_to_local(sp.convert_ST_datetime_to_object(job['completedOn']), "Australia/Sydney").strftime("%m/%d/%Y") if job['completedOn'] is not None else "None"
#         formatted['Job #'] = job['jobNumber'] if job['jobNumber'] is not None else "None"
#         formatted['invoiceId'] = job['invoiceId'] if job['invoiceId'] is not None else "None"
#         return formatted

#     def format_invoice(invoice):
#         formatted = {}
#         formatted['Suburb'] = invoice['customerAddress']['city']
#         formatted['Jobs Subtotal'] = float(invoice['subTotal'])
#         formatted['Balance'] = float(invoice['balance'])
#         formatted['Costs'] = get_job_cost(invoice)
#         formatted['Commissionable Sales'] = round(formatted['Jobs Subtotal'] - formatted['Costs'],2)
#         formatted['invoiceId'] = invoice['id']
#         return formatted

#     def format_payment(payment):
#         output = []
#         for invoice in payment['appliedTo']:
#             formatted = {}
#             formatted['invoiceId'] = invoice['appliedTo']
#             formatted['Payment Types'] = payment['type']
#             output.append(formatted)

#         return output

#     def get_invoice_ids(job_response):
#         return [str(job['invoiceId']) for job in job_response]
    
#     check_and_update_ss_for_data_service(state)

#     start_time = datetime.combine(start_date, time(0,0,0))
#     end_time = datetime.combine(end_date, time(23,59,59))
 
#     technicians_response = ss[f'st_data_service_{state}'].get_all_technicians()
#     technicians = format_employee_list(technicians_response)

#     job_response = ss[f'st_data_service_{state}'].get_jobs_created_between(start_time, end_time)
#     invoice_ids = get_invoice_ids(job_response)

#     unsuccessful_tag = get_unsuccessful_tag()
#     jobs = [format_job(job, technicians, unsuccessful_tag) for job in job_response]
#     jobs_df = pd.DataFrame(jobs)

#     del job_response
#     del jobs

#     invoice_response = ss[f'st_data_service_{state}'].get_invoices_by_id(invoice_ids)
#     invoices = [format_invoice(invoice) for invoice in invoice_response]
#     invoices_df = pd.DataFrame(invoices)

#     del invoice_response
#     del invoices

#     jobs = pd.merge(jobs_df, invoices_df, on='invoiceId', how='left')
#     jobs = jobs.loc[:, ['Sold By', 'Created Date', 'Job #', 'Suburb', 'Jobs Subtotal', 'Status', 'Completion Date', 'Commissionable Sales']]
#     jobs = jobs.sort_values(by='Status')
#     jobs['completed'] = jobs['Status'] == 'Completed'
    
#     unsuccessful_jobs = jobs[jobs['Status'] == "Unsuccessful"]
#     successful_jobs = jobs[jobs['Status'] != "Unsuccessful"]



#     commission_summary = successful_jobs.groupby(by=['Sold By'], as_index=False).agg({
#         'Commissionable Sales': 'sum',
#         'completed': 'all'
#     })

#     commission_summary.rename(columns={'completed': 'All Jobs Completed?'}, inplace=True)

#     # Create an Excel file in memory
#     buffer = BytesIO()
#     with pd.ExcelWriter(buffer) as writer:
#         commission_summary.to_excel(writer, sheet_name="Summary", index=False)
        
#         techs = commission_summary['Sold By'].to_list()
#         for tech in techs:
#             tech_successful = successful_jobs[successful_jobs['Sold By'] == tech]
#             tech_unsuccessful = unsuccessful_jobs[unsuccessful_jobs['Sold By'] == tech]
#             tech_full = pd.concat([tech_successful, tech_unsuccessful], ignore_index=True)
#             tech_full.to_excel(writer, sheet_name=tech, index=False)

#     # successful_jobs = pd.merge(successful_jobs, invoices_df, on='invoiceId', how='left')
#     # successful_jobs_income = successful_jobs[successful_jobs['Jobs Subtotal'] > 0]

#     return buffer

def get_timesheets_for_tech(tech_id, state, start_date, end_date):
    
    def format_timesheet(timesheet):
        formatted = {}
        formatted['id'] = timesheet['id']
        formatted['technicianId'] = timesheet['technicianId']
        formatted['jobId'] = timesheet['jobId']
        formatted['appointmentId'] = timesheet['appointmentId']
        formatted['dispatchedOn'] = sp.convert_ST_datetime_to_local_obj(timesheet['dispatchedOn'], ss[f'st_data_service_{state}'].timezone) if timesheet['dispatchedOn'] is not None else None
        formatted['arrivedOn'] = sp.convert_ST_datetime_to_local_obj(timesheet['arrivedOn'], ss[f'st_data_service_{state}'].timezone) if timesheet['arrivedOn'] is not None else None
        formatted['doneOn'] = sp.convert_ST_datetime_to_local_obj(timesheet['doneOn'], ss[f'st_data_service_{state}'].timezone) if timesheet['doneOn'] is not None else None
        return formatted


    check_and_update_ss_for_data_service(state)

    start_time = datetime.combine(start_date, time(0,0,0))
    end_time = datetime.combine(end_date, time(23,59,59))

    timesheet_data = ss[f'st_data_service_{state}'].get_api_data_between('payroll', 'jobs/timesheets', start_time, end_time, 'created')
    # timesheet_data = ss[f'st_data_service_{state}'].get_api_data_between('payroll', 'jobs/timesheets', start_time, end_time, 'created', options={'technicianId': tech_id})
    timesheet_data = [format_timesheet(timesheet) for timesheet in timesheet_data]
    
    timesheet_data = pd.DataFrame(timesheet_data).sort_values(by=['arrivedOn'])

    timesheet_data['arrivedOn'] = timesheet_data['arrivedOn'].dt.round('15min')
    timesheet_data['doneOn'] = timesheet_data['doneOn'].dt.round('15min')
    timesheet_data['arrivedDate'] = timesheet_data['arrivedOn'].dt.date
    timesheet_data['arrivedTime'] = timesheet_data['arrivedOn'].dt.time
    timesheet_data['doneTime'] = timesheet_data['doneOn'].dt.time
    first_rows = timesheet_data.groupby(['technicianId', 'arrivedDate'], as_index=False).first()[['technicianId', 'arrivedDate', 'arrivedTime', 'jobId']]
    last_rows = timesheet_data.groupby(['technicianId', 'arrivedDate'], as_index=False).last()[['technicianId', 'arrivedDate', 'doneTime', 'jobId']]
    timesheet_data = pd.merge(first_rows, last_rows, left_on=['technicianId', 'arrivedDate'], right_on=['technicianId', 'arrivedDate'])

    return timesheet_data


def update_job_external_data(job_id, state, data):
    # data is just a dict with keys and values, it gets converted here.
    external_data_payload = [
        {"key": key, "value": value}
        for key, value in data.items()
    ]

    check_and_update_ss_for_data_service(state)
    print(external_data_payload)
    ss[f'st_data_service_{state}'].patch_job_external_data(job_id, external_data_payload, ss.app_guid, patch_mode="Replace")
