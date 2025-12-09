from dataclasses import dataclass, field, fields
import pandas as pd
from datetime import date, time, datetime
from pprint import pprint



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
    date: date
    suburb: str
    subtotal: float
    materials: float
    merchant_fees: float
    profit: float
    payment_types: list = field(compare=False)
    eftpos: float
    cash: float
    payment_plan: float
    category: str

    def diff(self, value: 'JobData'):
        diff = {}
        for field in fields(self):
            self_val = getattr(self, field.name)
            other_val = getattr(value, field.name)
            if self_val != other_val:
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

def extract_jobs_from_week(df: pd.DataFrame, week_ranges: dict, week_end_date: date) -> WeekData:
    """
    Gets job data from the week specified.
    """
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
            j_date = row_data.iloc[2]
            j_num = row_data.iloc[3]
            j_suburb = row_data.iloc[4]
            j_subtotal = row_data.iloc[5]
            j_materials = row_data.iloc[6]
            j_merchantfee = row_data.iloc[7]
            j_profit = row_data.iloc[8]
            j_paymenttypes = row_data.iloc[9].split(', ') if type(row_data.iloc[9]) == str else ['N/A']
            j_eftpos = row_data.iloc[21]
            j_cash = row_data.iloc[22]
            j_paymentplan = row_data.iloc[23]
            
            job = JobData(j_num, j_date, j_suburb, j_subtotal, j_materials, j_merchantfee, j_profit, j_paymenttypes, j_eftpos, j_cash, j_paymentplan, curr_cat)
            jobs[j_num] = job

        if type(row_data.iloc[1]) == str:
            curr_cat = row_data.iloc[1]

    week_data = WeekData(jobs)
    return week_data


def flatten_list(nested_list):
    return [item for sublist in nested_list for item in sublist]

def main():
    MANUAL_EXCEL_FILE = '/Users/albie/Documents/code/github repos/TitanReporting/commission_tester/data/FY2026 - VIC Commission Sheet.xlsx'
    AUTO_EXCEL_FILE = '/Users/albie/Documents/code/github repos/TitanReporting/commission_tester/data/test_sheet.xlsx'
    MANUAL_SHEET_NAME = 'Bradley'
    STATE = 'VIC'
    AUTO_SHEET_NAME = f'{MANUAL_SHEET_NAME} {STATE}'
    TEST_DATE = datetime(2025,11,30,0,0)

    auto_excel = pd.ExcelFile(AUTO_EXCEL_FILE)
    manual_excel = pd.ExcelFile(MANUAL_EXCEL_FILE)

    manual_df = manual_excel.parse(sheet_name=MANUAL_SHEET_NAME)
    auto_df = auto_excel.parse(sheet_name=AUTO_SHEET_NAME, header=None)
    # print(auto_df)

    manual_week_ranges = get_week_data_ranges(manual_df)
    auto_week_ranges = get_week_data_ranges(auto_df)
    # print(auto_week_ranges)

    manual_test_week = extract_jobs_from_week(manual_df, manual_week_ranges, TEST_DATE)
    auto_test_week = extract_jobs_from_week(auto_df, auto_week_ranges, TEST_DATE)
    
    # manual_jobs = flatten_list([job.split('/') if type(job)==str else [job] for job in manual_test_week.jobs.keys()])
    # auto_jobs = flatten_list([job.split('/') if type(job)==str else [job] for job in auto_test_week.jobs.keys()])
    # print(manual_jobs)
    # print(auto_jobs)

    manual_jobs = [int(j) if type(j) != float else 0 for j in flatten_list([job.split('/') if type(job)==str else [job] for job in manual_test_week.jobs.keys()])]
    auto_jobs = [int(j) if type(j) != float else 0 for j in flatten_list([job.split('/') if type(job)==str else [job] for job in auto_test_week.jobs.keys()])]

    manual_jobs_set = set(manual_jobs)
    auto_jobs_set = set(auto_jobs)

    in_man_not_auto = manual_jobs_set.difference(auto_jobs_set)
    in_auto_not_man = auto_jobs_set.difference(manual_jobs_set)

    print(in_man_not_auto)
    print(in_auto_not_man)

    # j1 = JobData(1234, date(2025,11,24), 'Mascot', 15564.2, 1235.1, 0, 3000.0, ['Credit Card', 'EFT'], 15564.2*1.1, 0, 0, "COMPLETED & PAID JOBS")
    # j2 = JobData(123, date(2025,11,25), 'Mascot', 1234.2, 123.1, 0, 1000.0, ['Cash'], 0, 1234.2*1.1, 0, "COMPLETED & PAID JOBS")
    # j3 = JobData(1234, date(2025,11,24), 'Mascot', 15564.2, 1235.1, 0, 3000.0, ['Credit Card', 'EFT'], 15564.2*1.1, 0, 0, "COMPLETED & PAID JOBS")

    # # w1 = WeekData({
    # #     j1.num: j1,
    # #     j2.num: j2
    # # })

    # # print(j1)
    # # print(j2)
    # # print(w1)
    # pprint(j1.diff(j2))
    # pprint(j1.diff(j3))
    # # week_ranges = get_week_data_ranges(sheet)
    # # test_week = list(week_ranges.keys())[1]
    # # print(test_week)
    # # pprint(extract_jobs_from_week(sheet, week_ranges, test_week))

if __name__ == '__main__':
    main()
