from __future__ import annotations

from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json

from servicetitan_api_client import ServiceTitanClient
import modules.lookup_tables as lookup
import modules.helpers as helpers

def format_job(job, client: ServiceTitanClient, tenant_tags: list, exdata_key='docchecks_testing'):
    
    def check_unsuccessful(job, tags):
        unsuccessful_tags = {tag.get("id") for tag in tags if "Unsuccessful" in tag.get("name")}
        job_tags = set(job.get('tagTypeIds'))
        return bool(unsuccessful_tags & job_tags)
    
    # if 116255355 in job['tagTypeIds'] or 
    if job['jobStatus'] == 'Canceled': 
        return None
    formatted = {}
    if job['soldById'] is not None:
        formatted['sold_by'] = str(job['soldById'])

    else:
        if job['appointmentCount'] != job['num_of_appts_in_mem']:
            # print(f"Appt # not matching, {job['id']}, in mem: {job['num_of_appts_in_mem']}, in job: {job['appointmentCount']}")
            url = client.build_url("dispatch", "appointment-assignments")
            appt_assmnts = client.get_all(url, params={'jobId': job['id']})
            appt_assmnts = [format_appt_assmt(appt) for appt in appt_assmnts]
            job_appt_techs = {appt['tech_id'] for appt in appt_assmnts}
        else:
            job_appt_techs = job['appt_techs']
        if len(job_appt_techs) == 1:
            formatted['sold_by'] = list(job_appt_techs)[0]
        elif len(job_appt_techs) == 0:
            formatted['sold_by'] = 'Manual Check'
        else:
            # just add to manual check list for now, might separate to different check list in future.
            formatted['sold_by'] = 'Manual Check'

    formatted['created_str'] = client.st_date_to_local(job['createdOn'], fmt="%d/%m/%Y")
    formatted['created_dt'] = client.from_utc(job['createdOn'])
    formatted['completed_str'] = client.st_date_to_local(job['completedOn'], fmt="%m/%d/%Y") if job['completedOn'] is not None else "No data"
    formatted['num'] = job['jobNumber'] if job['jobNumber'] is not None else -1
    formatted['status'] = job['jobStatus'] if job['jobStatus'] is not None else "No data"
    formatted['invoiceId'] = job['invoiceId'] if job['invoiceId'] is not None else -1
    externalData = get_external_data_by_key(job['externalData'], key=exdata_key)
    formatted.update(format_external_data_for_xl(externalData))
    formatted['unsuccessful'] = check_unsuccessful(job, tenant_tags)
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
    if invoice['items']:
        formatted['summary'] = '\n'.join([i['description'].split('<')[0] for i in invoice['items']])
    else:
        formatted['summary'] = 'no invoice items'
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

def format_appt_assmt(appt):
    formatted = {}
    formatted['job_id'] = appt['jobId']
    formatted['tech_id'] = appt['technicianId']
    formatted['tech_name'] = appt['technicianName']
    return formatted

def format_appt(appt): # TODO: finish figuring out what "start time" counts.
    formatted = {}
    formatted['job_id'] = appt['jobId']
    formatted['appt_start'] = appt['start']
    formatted['appt_end'] = appt['end']
    return formatted

# def get_sales_and_installer_codes(roles_reponse):
#     role_codes = {
#         's': [],
#         'i': []
#     }
#     for role in roles_reponse:
#         if role['name'] == 'Technician - Sales':
#             role_codes['s'].append(role['id'])
#         if role['name'] == 'Technician - Installer':
#             role_codes['i'].append(role['id'])
#     return role_codes

def format_employee_list(employee_response):
    # input can be either technician response or employee response

    # def test_role(emp_roles, test_roles):
    #     return bool(set(emp_roles) & set(test_roles))

    formatted = {}
    for employee in employee_response:
        team = employee.get('team', 'O')
        if team:
            formatted[employee['id']] = {'name': employee['name'], 'team': team[0]}
            formatted[employee['userId']] = {'name': employee['name'], 'team': team[0]}
        else:
            formatted[employee['id']] = {'name': employee['name'], 'team': 'O'}
            formatted[employee['userId']] = {'name': employee['name'], 'team': 'O'}
    return formatted

# def format_employee_list(employee_response, sales_codes=None):
#     # input can be either technician response or employee response

#     def test_sales(emp_roles, sales_roles):
#         return bool(set(emp_roles) & set(sales_roles))

#     formatted = {}
#     sales = set()
#     for employee in employee_response:
#         if sales_codes:
#             test = test_sales(employee['roleIds'], sales_codes)
#             formatted[employee['id']] = {'name': employee['name'], 'sales': test}
#             formatted[employee['userId']] = {'name': employee['name'], 'sales': test}
#             if test:
#                 sales.add(employee['id'])
#                 sales.add(employee['userId'])
#         else:   
#             formatted[employee['id']] = {'name': employee['name'], 'sales': False}
#             formatted[employee['userId']] = {'name': employee['name'], 'sales': False}
#     return formatted, sales

def group_jobs_by_tech(job_records, employee_map):
    jobs_by_tech: dict[str, list[dict]] = {}
    for j in job_records:
        tid = j.get("sold_by")
        if not tid:
            continue
        if tid == 'Manual Check':
            name = tid + 'O'
        else:
            tech_info = employee_map.get(int(tid))
            name = tech_info.get("name", f"{tid}")
            tech_role = tech_info.get('team', 'O')
            name = name + tech_role
        j_category = helpers.categorise_job(j)
        jobs_by_tech.setdefault(name, dict()).setdefault(j_category, []).append(j)
    return jobs_by_tech

def group_appt_assmnts_by_job(appt_assmnts):
    appt_assmnts_by_job: dict[str, list] = {}
    for a in appt_assmnts:
        job_id = a.get("job_id")
        if not job_id: 
            continue
        appt_assmnts_by_job.setdefault(job_id, list()).append(a.get("tech_id"))
    num_appts_per_job = {job_id: len(appts) for job_id, appts in appt_assmnts_by_job.items()}
    return appt_assmnts_by_job, num_appts_per_job