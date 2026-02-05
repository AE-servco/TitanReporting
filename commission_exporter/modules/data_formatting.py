from __future__ import annotations

from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json

from servicetitan_api_client import ServiceTitanClient
import modules.lookup_tables as lookup
import modules.helpers as helpers

def check_unsuccessful(job, tags):
    unsuccessful_tags = {tag.get("id") for tag in tags if "Unsuccessful" in tag.get("name") or "Cancelled" in tag.get("name")}
    job_tags = set(job.get('tagTypeIds'))
    return bool(unsuccessful_tags & job_tags)

def check_complaint(job, tags):
    complaint_tags = {tag.get("id") for tag in tags if "complaint" in tag.get("name").lower()}
    job_tags = set(job.get('tagTypeIds'))
    return bool(complaint_tags & job_tags)

def format_job(job, client: ServiceTitanClient, tenant_tags: list, exdata_key='docchecks_live'):
    
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
            formatted['sold_by'] = str(list(job_appt_techs)[0])
        elif len(job_appt_techs) == 0:
            formatted['sold_by'] = 'Manual Check'
        else:
            # just add to manual check list for now, might separate to different check list in future.
            formatted['sold_by'] = str(list(job_appt_techs)[0])
            # formatted['sold_by'] = 'Manual Check'

    if job['first_appt']:
        formatted['first_appt_start_dt'] = client.from_utc(job['first_appt']['start'])
        formatted['first_appt_start_str'] = client.st_date_to_local((job['first_appt']['start']), fmt="%d/%m/%Y")

    formatted['job_id'] = job['id']
    formatted['created_str'] = client.st_date_to_local(job['createdOn'], fmt="%d/%m/%Y")
    formatted['created_dt'] = client.from_utc(job['createdOn'])
    formatted['completed_str'] = client.st_date_to_local(job['completedOn'], fmt="%d/%m/%Y") if job['completedOn'] is not None else "No data"
    formatted['completed_dt'] = client.from_utc(job['completedOn']) if job['completedOn'] is not None else None
    formatted['num'] = job['jobNumber'] if job['jobNumber'] is not None else -1
    formatted['status'] = job['jobStatus'] if job['jobStatus'] is not None else "No data"
    formatted['invoiceId'] = job['invoiceId'] if job['invoiceId'] is not None else -1
    externalData = get_external_data_by_key(job['externalData'], key=exdata_key)
    formatted.update(format_external_data_for_xl(externalData))
    formatted['unsuccessful'] = check_unsuccessful(job, tenant_tags)
    formatted['complaint_tag_present'] = check_complaint(job, tenant_tags)
    formatted['total'] = job.get('total')
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
        output = {check_map[k]: v for k,v in exdata.items()}
        if not output['Invoice Signed']:
            if output.get('Invoice Not Signed (Client Offsite)'):
                output['Invoice Signed'] = 1
        return output
    return {v: 0 for k,v in check_map.items()}

def get_invoice_ids(job_response) -> List[str]:
    return [str(job['invoiceId']) for job in job_response]

def get_job_ids(appt_response) -> List[str]:
    return [str(appt['jobId']) for appt in appt_response]

def format_invoice(invoice):
    formatted = {}
    formatted['suburb'] = invoice['customerAddress']['city']
    formatted['inv_subtotal_orig'] = float(invoice['subTotal'])
    formatted['inv_subtotal'] = formatted['inv_subtotal_orig'] 
    formatted['balance'] = float(invoice['balance'])
    formatted['amt_paid'] = round(float(invoice['total']) - float(invoice['balance']),2)
    formatted['invoiceId'] = invoice['id']
    if invoice['items']:
        formatted['summary'] = '\n'.join([i['description'].split('<')[0] for i in invoice['items']])
        for i in invoice['items']:
            if "cof" in i['skuName'].lower():
                formatted['inv_subtotal'] -= float(i['price'])

        # formatted['inv_items_skuids'] = ', '.join([str(i['skuId']) for i in invoice['items']])
        # formatted['inv_items_skuNames'] = ', '.join([i['skuName'] for i in invoice['items']])
        # formatted['inv_items_displayName'] = ', '.join([i['displayName'] for i in invoice['items']])
        # formatted['inv_items_ids'] = ', '.join([str(i['id']) for i in invoice['items']])
        # formatted['inv_items_cost'] = ', '.join([str(i['cost']) for i in invoice['items']])
        # formatted['inv_items_price'] = ', '.join([str(i['price']) for i in invoice['items']])
        # formatted['inv_items_total'] = ', '.join([str(i['total']) for i in invoice['items']])
    else:
        formatted['summary'] = 'no invoice items'
        # formatted['inv_items_skuids'] = ""
        # formatted['inv_items_skuNames'] = ""
        # formatted['inv_items_displayName'] = ""
        # formatted['inv_items_ids'] = ""
        # formatted['inv_items_cost'] = ""
        # formatted['inv_items_price'] = ""
        # formatted['inv_items_total'] = ""
    return formatted

def format_payment(payment, client: ServiceTitanClient):
    output = []
    for invoice in payment['appliedTo']:
        formatted = {}
        formatted['invoiceId'] = invoice['appliedTo']
        formatted['payment_types'] = payment['type']
        try:    
            formatted['payment_dates'] = client.st_date_to_local(payment['date'])[:10] # cut off at just date
            # formatted['payment_dates'] = client.st_date_to_local(invoice['appliedOn'])[:10] # cut off at just date
        except KeyError:
            formatted['payment_dates'] = 'no payment date'
        all_payment_types = lookup.get_all_payment_types()
        formatted['payment_details'] = f"{all_payment_types.get(payment['type'], payment['type'])}|{invoice.get('appliedAmount', '0')}"
        # formatted['payment_details'] = {'type': payment['type'], 'amount': invoice.get('appliedAmount', '0')}
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

# def get_first_appts(appts) -> Dict[int, Dict]:
#     output = {}
#     for appt in appts:
#         if appt.get('appointmentNumber', '').endswith('-1'):
#             output.setdefault(appt['jobId'], []).append(appt)
#     return output

def get_first_appts(appts) -> List:
    return [appt for appt in appts if appt.get('appointmentNumber', '').endswith('-1')]

def extract_id_to_key(input: List, id_val: str = 'id', keep_id: bool = True) -> Dict:
    output = {}
    for item in input:
        value_data = item.copy()
        if not keep_id:
            del value_data[id_val]
        output[item[id_val]] = value_data
    return output

def format_estimate(estimate, sold: bool): # TODO: finish figuring out what "start time" counts.
    formatted = {}
    try:
        if sold:
            if estimate['status']['name'] != "Sold":
                return None
        else:
            if estimate['status']['name'] == "Sold":
                return None
    except KeyError:
        return None
    formatted['job_id'] = estimate['jobId']
    formatted['est_subtotal'] = estimate['subtotal']
    return formatted


def format_employee_list(employee_response):
    # input can be either technician response or employee response
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

def group_jobs_by_tech(job_records, employee_map, end_date, relevant_holidays) -> dict:
    jobs_by_tech: dict[str, list[dict]] = {}
    for j in job_records:
        tid = j.get("sold_by")
        # print(tid)
        if not tid:
            continue
        if tid == 'Manual Check':
            name = tid + 'O'
        else:
            tech_info = employee_map.get(int(tid))
            # print(tech_info)
            if tech_info is None:
                name = 'Manual Check' + 'O'
            else:
                name = tech_info.get("name", f"{tid}")
                tech_role = tech_info.get('team', 'O')
                name = name + tech_role
        # print(name)
        j_category = helpers.categorise_job(j, end_date, relevant_holidays)
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