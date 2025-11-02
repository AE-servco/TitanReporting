import pandas as pd
from google.cloud import secretmanager
from datetime import datetime, time, date
from zoneinfo import ZoneInfo
import pytz

import servicepytan as sp

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

def get_data_service(state):
    state_code = state_codes()[state]
    st_conn = sp.auth.servicepytan_connect(app_key=get_secret("ST_app_key_tester"), tenant_id=get_secret(f"ST_tenant_id_{state_code}"), client_id=get_secret(f"ST_client_id_{state_code}"), 
    client_secret=get_secret(f"ST_client_secret_{state_code}"), timezone="Australia/Sydney")
    st_data_service = sp.DataService(conn=st_conn)

    return st_data_service

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

    st_data_service = get_data_service(state)

    start_time = datetime.combine(start_date, time(0,0,0))
    end_time = datetime.combine(end_date, time(23,59,59))

    invoice_response = st_data_service.get_invoices_between(start_time, end_time)

    invoices = [format_invoice(invoice) for invoice in invoice_response]
    
    return pd.DataFrame(invoices)

def convert_df_for_download(df):
    if df is None:
        df = pd.DataFrame()
    return df.to_csv(index=False).encode("utf-8")

def get_commission_data(state, start_date, end_date):

    unsucessfultag = 116255355
    # Filter for unsuccessful tag on job, soldBy should be in job, get technician name through settings endpoint
    # Payment type???
    # Do they get comms if its just the call out fee?

    def format_technicians_list(technicians_response):
        formatted = {}
        for technician in technicians_response:
            formatted[technician['id']] = technician['name']
        return formatted

    def format_job(job, technicians):
        if 116255355 in job['tagTypeIds'] or job['jobStatus'] == 'Canceled': # Unsuccessful or cancelled 
            return None
        formatted = {}
        if job['soldById'] is not None:
            formatted['Sold By'] = technicians[job['soldById']]
        else:
            appts = st_data_service.get_appointment_assignments_by_job_id(job['id'])
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

    st_data_service = get_data_service(state)

    start_time = datetime.combine(start_date, time(0,0,0))
    end_time = datetime.combine(end_date, time(23,59,59))

    technicians_response = st_data_service.get_all_technicians()
    technicians = format_technicians_list(technicians_response)

    job_response = st_data_service.get_jobs_created_between(start_time, end_time)
    jobs_w_nones = [format_job(job, technicians) for job in job_response]
    jobs = [job for job in jobs_w_nones if job is not None]
    jobs_df = pd.DataFrame(jobs)
    
    invoice_ids = get_invoice_ids(job_response)
    invoice_response = st_data_service.get_invoices_by_id(invoice_ids)
    invoices = [format_invoice(invoice) for invoice in invoice_response]
    invoices_df = pd.DataFrame(invoices)

    payments_response = st_data_service.get_payments_for_invoices(invoice_ids)
    payments = [format_payment(payment) for payment in payments_response]
    payments_flat = flatten_list(payments)
    payments_df = pd.DataFrame(payments_flat)
    payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(','.join)

    merged = pd.merge(pd.merge(jobs_df, invoices_df, on='invoiceId', how='left'), payments_grouped, on='invoiceId', how='left')

    merged_reordered = merged.loc[:, ['Sold By', 'Created Date', 'Job #', 'Suburb', 'Jobs Subtotal', 'Payment Types', 'Status', 'Completion Date', 'Payments']]

    return merged_reordered.sort_values(by=['Sold By', 'Status'])

def get_doc_check_checker_data(state):

    # def format_invoice(invoice):
    #     formatted = {}
    #     invoice_num = invoice['referenceNumber']
    #     if state == "WA" and invoice_num.startswith("1"):
    #         formatted['*InvoiceNumber'] = 'W' + invoice_num
    #     else:
    #         formatted['*InvoiceNumber'] = invoice_num
    #     formatted['Invoice Date'] = datetime.fromisoformat(invoice['invoiceDate'].replace('Z', '+00:00')).strftime("%m/%d/%Y")
    #     formatted['*ContactName'] = invoice['customer']['name']
    #     formatted['Location Address'] = format_address(invoice['customerAddress'])
    #     formatted['*Description'] = invoice['job']['type']
    #     formatted['*UnitAmount'] = invoice['subTotal']
    #     formatted['*TaxType'] = "GST on Income"
    #     formatted['Sum'] = invoice['total']
    #     formatted['*AccountCode'] = account_codes()[state]
    #     return formatted

    # st_data_service = get_data_service(state)

    # invoice_response = st_data_service.get_invoices_between(start_time, end_time)

    # invoices = [format_invoice(invoice) for invoice in invoice_response]
    
    # return pd.DataFrame(invoices)
    return