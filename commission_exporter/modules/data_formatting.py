from __future__ import annotations

from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json

from servicetitan_api_client import ServiceTitanClient
import modules.lookup_tables as lookup

def format_job(job, client: ServiceTitanClient, tech_sales: list, exdata_key='docchecks'):
    # if 116255355 in job['tagTypeIds'] or 
    if job['jobStatus'] == 'Canceled': 
        return None
    formatted = {}
    if job['soldById'] is not None:
        formatted['sold_by'] = str(job['soldById'])
    elif 116255355 in job['tagTypeIds']:
        formatted['sold_by'] = 'No data - unsuccessful'
    else:
        url = client.build_url("dispatch", "appointment-assignments")
        appts = client.get_all(url, params={'jobId': job['id']})
        if appts:
            # appt_techs = [appt['technicianId'] for appt in appts]
            sales_techs_on_job = set()
            for appt in appts:
                if appt['technicianId'] in tech_sales:
                    sales_techs_on_job.add(appt['technicianId'])
            if len(sales_techs_on_job) == 0:
                formatted['sold_by'] = 'No Sales Plumber'
            elif len(sales_techs_on_job) == 1:
                formatted['sold_by'] = str(list(sales_techs_on_job)[0])
            else:
                formatted['sold_by'] = ', '.join([str(tech) for tech in list(sales_techs_on_job).sort()])
        else:
            formatted['sold_by'] = '-1'

    formatted['created_str'] = client.st_date_to_local(job['createdOn'], fmt="%d/%m/%Y")
    formatted['created_dt'] = client.from_utc(job['createdOn'])
    formatted['completed_str'] = client.st_date_to_local(job['completedOn'], fmt="%m/%d/%Y") if job['completedOn'] is not None else "No data"
    formatted['num'] = job['jobNumber'] if job['jobNumber'] is not None else -1
    formatted['status'] = job['jobStatus'] if job['jobStatus'] is not None else "No data"
    formatted['invoiceId'] = job['invoiceId'] if job['invoiceId'] is not None else -1
    externalData = get_external_data_by_key(job['externalData'], key=exdata_key)
    formatted.update(format_external_data_for_xl(externalData))
    formatted['unsuccessful'] = 116255355 in job['tagTypeIds']
    return formatted

def get_external_data_by_key(data, key='docchecks'):
    if data is None:
        return None
    for d in data:
        if d['key'] == key:
            return json.loads(d['value'])
    return None

def format_external_data_for_xl(exdata):
    check_map = lookup.get_doc_check_criteria()
    if exdata:
        return {check_map[k]: v for k,v in exdata.items()}
    return {v: 0 for k,v in check_map.items()}

def get_invoice_ids(job_response):
    return [str(job['invoiceId']) for job in job_response]

def format_invoice(invoice):
    formatted = {}
    formatted['suburb'] = invoice['customerAddress']['city']
    formatted['subtotal'] = float(invoice['subTotal'])
    formatted['balance'] = float(invoice['balance'])
    formatted['amt_paid'] = round(float(invoice['total']) - float(invoice['balance']),2)
    formatted['invoiceId'] = invoice['id']
    return formatted

def format_payment(payment):
    output = []
    for invoice in payment['appliedTo']:
        formatted = {}
        formatted['invoiceId'] = invoice['appliedTo']
        formatted['payment_types'] = payment['type']
        formatted['payment_amt'] = payment['type'][:2] + invoice.get('appliedAmount', '0')
        output.append(formatted)
    return output

def format_employee_list(employee_response, sales_codes=None):
    # input can be either technician response or employee response

    def test_sales(emp_roles, sales_roles):
        return bool(set(emp_roles) & set(sales_roles))

    formatted = {}
    sales = set()
    for employee in employee_response:
        if sales_codes:
            test = test_sales(employee['roleIds'], sales_codes)
            formatted[employee['id']] = {'name': employee['name'], 'sales': test}
            formatted[employee['userId']] = {'name': employee['name'], 'sales': test}
            if test:
                sales.add(employee['id'])
                sales.add(employee['userId'])
        else:   
            formatted[employee['id']] = {'name': employee['name'], 'sales': False}
            formatted[employee['userId']] = {'name': employee['name'], 'sales': False}
    return formatted, sales