from servicetitan_api_client import ServiceTitanClient

def format_employee_list(employee_response):
    # input can be either technician response or employee response
    formatted = {}
    for employee in employee_response:
        formatted[employee['id']] = employee['name']
        formatted[employee['userId']] = employee['name']
    return formatted

def add_appt_info(job, appt, modifier='first'):
    job[f'{modifier}_appt_start'] = appt.get("start")
    job[f'{modifier}_appt_end'] = appt.get("end")
    job[f'{modifier}_appt_arrival_start'] = appt.get("arrivalWindowStart")
    job[f'{modifier}_appt_arrival_end'] = appt.get("arrivalWindowEnd")
    job[f'{modifier}_appt_num'] = appt.get("appointmentNumber")
    return job

def get_invoice_ids(job_response):
    return [str(job['invoiceId']) for job in job_response]

def format_invoice(invoice):
    formatted = {}
    formatted['suburb'] = invoice['customerAddress']['city']
    formatted['subtotal'] = float(invoice['subTotal'])
    formatted['balance'] = float(invoice['balance'])
    formatted['amt_paid'] = round(float(invoice['total']) - float(invoice['balance']),2)
    formatted['invoiceId'] = invoice['id']
    formatted['sent_status'] = invoice['sentStatus']
    if invoice['items']:
        desc_list = []
        for i in invoice['items']:
            spl = i['description'].split('<')
            if spl[0] == '':
                if spl[1][1] == '>':
                    desc_list.append(spl[1][2:])
                else:
                    desc_list.append(spl[1])
            else:
                desc_list.append(spl[0])
        formatted['summary'] = '|'.join(desc_list)
    else:
        formatted['summary'] = 'no invoice items'
    return formatted

def format_payments(payments):
    output = {}
    for payment in payments:
        for invoice in payment.get('appliedTo', []):
            inv_dict = {}
            inv_dict['payment_type'] = payment['type']
            inv_dict['payment_amt'] = invoice.get('appliedAmount', '0')
            output.setdefault(invoice['appliedTo'], []).append(inv_dict)
    return output

# def format_job(job, client: ServiceTitanClient, tenant_tags: list, exdata_key='docchecks_live'):
    
#     def check_unsuccessful(job, tags):
#         unsuccessful_tags = {tag.get("id") for tag in tags if "Unsuccessful" in tag.get("name")}
#         job_tags = set(job.get('tagTypeIds'))
#         return bool(unsuccessful_tags & job_tags)
 
#     formatted = {}
#     formatted['created_str'] = client.st_date_to_local(job['createdOn'], fmt="%d/%m/%Y")
#     formatted['created_dt'] = client.from_utc(job['createdOn'])
#     formatted['completed_str'] = client.st_date_to_local(job['completedOn'], fmt="%m/%d/%Y") if job['completedOn'] is not None else "No data"
#     formatted['num'] = job['jobNumber'] if job['jobNumber'] is not None else -1
#     formatted['status'] = job['jobStatus'] if job['jobStatus'] is not None else "No data"
#     formatted['invoiceId'] = job['invoiceId'] if job['invoiceId'] is not None else -1
#     formatted['project_id'] = job['projectId'] if job['projectId'] is not None else -1
#     formatted['unsuccessful'] = check_unsuccessful(job, tenant_tags)
#     return formatted

def combine_job_data(jobs: list, invoices: dict, payments: dict):
    """Combines job, invoice, payment, and any other data for jobs into a single dictionary.
    
    Args:
        jobs: list of job data ServiceTitan API responses
        invoice: dict with invoice ids as keys and rest of invoice data as value
        payments: dict with invoice ids as keys and rest of payment data as value
    """

    for job in jobs:
        job['invoice_data'] = invoices.get(job['invoiceId'], {})
        job['payment_data'] = payments.get(job['invoiceId'], {})

    return jobs


