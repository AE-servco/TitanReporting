import pandas as pd
from google.cloud import secretmanager
from datetime import datetime, time, date
from zoneinfo import ZoneInfo
import pytz

import servicepytan as sp

def get_secret(secret_id, project_id="prestigious-gcp", version_id="latest"):
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/{secret_id}/versions/{version_id}"
    response = client.access_secret_version(request={"name": name})
    secret_payload = response.payload.data.decode("UTF-8")
    return secret_payload

# def state_codes():
#     codes = {
#         'NSW_old': 'alphabravo',
#         'VIC_old': 'victortango',
#         'QLD_old': 'echozulu',
#         'NSW': 'foxtrotwhiskey',
#         'WA': 'sierradelta',
#         'QLD': 'bravogolf',
#     }
#     return codes

def get_invoices_for_xero(tenant_stripped, start_date, end_date):

    def account_codes():
        codes = {
            'foxtrotwhiskey': '210-2',
            'sierradelta': '210-1',
            'victortango': '210-1',
            'echozulu': '210-1',
            'bravogolf': '210-4',
            'mikeecho': '210-3',
        }
        return codes

    # def format_address(customerAddress):
    #     address = f"{customerAddress['street']}, {customerAddress['city']}, {customerAddress['state']} {customerAddress['zip']}, {customerAddress['country']}"
    #     if customerAddress['unit']:
    #         address = f"{customerAddress['unit']}/{address}"
    #     return address

    def format_invoice(invoice):
        formatted = {}
        invoice_num = invoice['referenceNumber']
        if tenant_stripped == "sierradelta" and invoice_num.startswith("1"):
            formatted['*InvoiceNumber'] = 'W' + invoice_num
        else:
            formatted['*InvoiceNumber'] = invoice_num
        formatted['Invoice Date'] = datetime.fromisoformat(invoice['invoiceDate'].replace('Z', '+00:00')).strftime("%m/%d/%Y")
        formatted['*ContactName'] = invoice['customer']['name']
        # formatted['Location Address'] = format_address(invoice['customerAddress'])
        formatted['POAddressLine1'] = invoice['customerAddress']['street']
        if invoice['customerAddress']['unit']:
            formatted['POAddressLine1'] = f"{invoice['customerAddress']['unit']}/{formatted['POAddressLine1']}"
        formatted['POCity'] = invoice['customerAddress']['city']
        formatted['PORegion'] = invoice['customerAddress']['state']
        formatted['POPostalCode'] = invoice['customerAddress']['zip']
        formatted['POCountry'] = invoice['customerAddress']['country']
        formatted['*Description'] = invoice['job']['type']
        formatted['*UnitAmount'] = invoice['subTotal']
        formatted['*TaxType'] = "GST on Income"
        formatted['Sum'] = invoice['total']
        formatted['*AccountCode'] = account_codes()[tenant_stripped]
        return formatted

    st_conn = sp.auth.servicepytan_connect(app_key=get_secret("ST_app_key_tester"), tenant_id=get_secret(f"ST_tenant_id_{tenant_stripped}"), client_id=get_secret(f"ST_client_id_{tenant_stripped}"), 
    client_secret=get_secret(f"ST_client_secret_{tenant_stripped}"), timezone="Australia/Sydney")
    st_data_service = sp.DataService(conn=st_conn)

    start_time = datetime.combine(start_date, time(0,0,0))
    end_time = datetime.combine(end_date, time(23,59,59))

    invoice_response = st_data_service.get_invoices_between(start_time, end_time)

    invoices = [format_invoice(invoice) for invoice in invoice_response]
    
    return pd.DataFrame(invoices)

def convert_df_for_download(df):
    if df is None:
        df = pd.DataFrame()
    return df.to_csv(index=False).encode("utf-8")