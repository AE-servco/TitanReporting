from dataclasses import dataclass, field, fields
import pandas as pd
from datetime import date, time, datetime
from pprint import pprint
import math
import json



# sheet = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=None)

# count = 0
# for row in sheet.iterrows():
#     print(row)
#     count += 1
#     if count == 30:
#         break

@dataclass(order=True)
class JobData():
    num: int | str
    date: 'date'
    suburb: str
    subtotal: float
    materials: float = field(compare=False)
    merchant_fees: float = field(compare=False)
    profit: float = field(compare=False)
    payment_types: list = field(compare=False)
    eftpos: float
    cash: float
    payment_plan: float
    category: str
    other_j_nums: list = field(compare=False)

    def diff(self, value: 'JobData'):
        diff = {}
        for field in fields(self):
            self_val = getattr(self, field.name)
            other_val = getattr(value, field.name)
            if self_val != other_val:
                if type(self_val) == datetime or type(self_val) == datetime:
                    self_val = self_val.strftime("%d/%m/%Y")
                    other_val = other_val.strftime("%d/%m/%Y")
                diff[field.name] = (self_val, other_val)
        return diff

class JobDiff():
    # Class for differences in jobs? 
    pass

@dataclass
class WeekData():
    jobs: dict[int, JobData]

def get_week_data_ranges(df: pd.DataFrame):
    """ 
    extract the start and end rows for each week. 
    Returns dict: {week_end_date: (start,end)}
    """
    week_ranges = {}
    curr_row = 1
    curr_week = None
    curr_start = 1
    # curr_end = 1
    for row in df.iterrows():
        # print(row)
        if type(row[1].iloc[1]) == str and 'WEEKLY COMMISSION' in row[1].iloc[1]:
            if curr_week:
                week_ranges[curr_week] = (curr_start, curr_row-1)
            curr_week = row[1].iloc[6]
            if type(curr_week) == str:
                curr_week = datetime.strptime(curr_week, "%d/%m/%Y")
            curr_start = curr_row
        curr_row += 1
    week_ranges[curr_week] = (curr_start, curr_row-1)
    return week_ranges

def extract_summary_from_week(df: pd.DataFrame, week_ranges: dict, week_end_date: date):
    """
    Gets summary data from the week specified.
    """
    pass

def extract_jobs_from_week(df: pd.DataFrame, week_ranges: dict, week_end_date: date, manual_or_auto: str = 'manual') -> WeekData:
    """
    Gets job data from the week specified.
    """
    payment_type_map = {
        'Credit Card': 'Card',
        'EFT/Bank Transfer': 'EFT',
        'Cash': 'Cash',
        'N/A': 'N/A'
    }

    job_section = df[week_ranges[week_end_date][0]:week_ranges[week_end_date][1]]
    
    curr_cat = 'COMPLETED & PAID JOBS'
    in_job_section = False
    jobs = {}
    for row in job_section.iterrows():
        row_data = row[1]
        if not in_job_section:
            if type(row_data.iloc[1]) == str and 'COMPLETED & PAID JOBS' in row_data.iloc[1]:
                in_job_section = True
            continue
        # print(type(row_data[1]), row_data[1])
        if type(row_data.iloc[1]) == int:
            j_date = datetime.strptime(row_data.iloc[2], "%d/%m/%Y") if type(row_data.iloc[2]) == str else row_data.iloc[2] 
            j_num = row_data.iloc[3]
            j_nums = []
            if type(j_num) == str:
                j_nums = [int(n) for n in j_num.split('/')]
                j_num = j_nums[0]
            j_suburb = row_data.iloc[4]
            j_subtotal = row_data.iloc[5] if not math.isnan(row_data.iloc[5]) else 0
            j_materials = row_data.iloc[6] if not math.isnan(row_data.iloc[6]) else 0
            j_merchantfee = row_data.iloc[7] if not math.isnan(row_data.iloc[7]) else 0
            j_profit = row_data.iloc[8] if not math.isnan(row_data.iloc[8]) else 0
            j_paymenttypes = row_data.iloc[9].split(', ') if type(row_data.iloc[9]) == str else ['N/A']
            j_paymenttypes = [payment_type_map[jpt] if jpt in payment_type_map else jpt for jpt in j_paymenttypes]
            offset = 0
            if manual_or_auto == 'auto':
                offset = 1
            j_eftpos = row_data.iloc[offset + 21] if not math.isnan(row_data.iloc[offset + 21]) else 0
            j_cash = row_data.iloc[offset + 22] if not math.isnan(row_data.iloc[offset + 22]) else 0
            j_paymentplan = row_data.iloc[offset + 23] if not math.isnan(row_data.iloc[offset + 23]) else 0
            
            # if j_nums:
            #     for j in j_nums:
            #         job = JobData(j, j_date, j_suburb, j_subtotal, j_materials, j_merchantfee, j_profit, j_paymenttypes, j_eftpos, j_cash, j_paymentplan, curr_cat)
            #         jobs[j_num] = job
            job = JobData(j_num, j_date, j_suburb, j_subtotal, j_materials, j_merchantfee, j_profit, j_paymenttypes, j_eftpos, j_cash, j_paymentplan, curr_cat, j_nums)
            jobs[j_num] = job

        if type(row_data.iloc[1]) == str:
            curr_cat = row_data.iloc[1]

    week_data = WeekData(jobs)
    return week_data


def flatten_list(nested_list):
    return [item for sublist in nested_list for item in sublist]

def main():
    output = {}

    MANUAL_EXCEL_FILE = '/Users/albie/Documents/code/github repos/TitanReporting/commission_tester/data/FY2026 - VIC Commission Sheet.xlsx'
    AUTO_EXCEL_FILE = '/Users/albie/Documents/code/github repos/TitanReporting/commission_tester/data/test_sheet.xlsx'
    MANUAL_SHEET_NAME = 'Liam'
    STATE = 'VIC'
    # AUTO_SHEET_NAME = f'Jarrod {STATE}'
    AUTO_SHEET_NAME = f'{MANUAL_SHEET_NAME} {STATE}'
    TEST_DATE = datetime(2025,11,30,0,0)

    auto_excel = pd.ExcelFile(AUTO_EXCEL_FILE)
    manual_excel = pd.ExcelFile(MANUAL_EXCEL_FILE)

    manual_df = manual_excel.parse(sheet_name=MANUAL_SHEET_NAME)
    auto_df = auto_excel.parse(sheet_name=AUTO_SHEET_NAME, header=None)
    # print(manual_df)
    # print(auto_df)

    manual_week_ranges = get_week_data_ranges(manual_df)
    auto_week_ranges = get_week_data_ranges(auto_df)
    # print(manual_week_ranges)
    # print(auto_week_ranges)

    manual_test_week = extract_jobs_from_week(manual_df, manual_week_ranges, TEST_DATE)
    auto_test_week = extract_jobs_from_week(auto_df, auto_week_ranges, TEST_DATE, manual_or_auto='auto')
    # pprint(manual_test_week)
    # pprint(auto_test_week)

    manual_jobs = [int(j) if type(j) != float else 0 for j in flatten_list([job.split('/') if type(job)==str else [job] for job in manual_test_week.jobs.keys()])]
    auto_jobs = [int(j) if type(j) != float else 0 for j in flatten_list([job.split('/') if type(job)==str else [job] for job in auto_test_week.jobs.keys()])]

    manual_jobs_set = set(manual_jobs)
    auto_jobs_set = set(auto_jobs)
    # print(manual_jobs_set)
    # print(auto_jobs_set)

    in_man_not_auto = manual_jobs_set.difference(auto_jobs_set)
    in_auto_not_man = auto_jobs_set.difference(manual_jobs_set)

    # print(in_man_not_auto)
    # print(in_auto_not_man)

    if in_man_not_auto:
        output['jobs_missing_from_automation'] = sorted(list(in_man_not_auto))
    if in_auto_not_man:
        output['jobs_missing_from_manual'] = sorted(list(in_auto_not_man))

    jobs_in_both = manual_jobs_set.intersection(auto_jobs_set)

    output['job_diffs (manual, auto)'] = {}
    for job in jobs_in_both:
        # print(job)
        if manual_test_week.jobs.get(job) != auto_test_week.jobs.get(job):
            output['job_diffs (manual, auto)'][job] = manual_test_week.jobs.get(job).diff(auto_test_week.jobs.get(job))

    with(open(f'data/{MANUAL_SHEET_NAME}_differences.json', 'w')) as f:
        json.dump(output, f)

if __name__ == '__main__':
    main()
