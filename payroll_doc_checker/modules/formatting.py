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










