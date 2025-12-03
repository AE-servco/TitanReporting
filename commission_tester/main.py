from dataclasses import dataclass, field, fields
import pandas as pd
from datetime import date, time, datetime
from pprint import pprint


EXCEL_FILE = '/Users/albie/Documents/code/github repos/TitanReporting/commission_tester/FY2026 - VIC Commission Sheet.xlsx'
SHEET_NAME = 'Bradley'

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

    def __eq__(self, value: 'JobData'):
        if self.num == value.num \
            and self.date == value.date \
            and self.suburb == value.suburb \
            and self.subtotal == value.subtotal \
            and self.materials == value.materials \
            and self.merchant_fees == value.merchant_fees \
            and self.profit == value.profit \
            and self.payment_types == value.payment_types \
            and self.eftpos == value.eftpos \
            and self.cash == value.cash \
            and self.payment_plan == value.payment_plan \
            and self.category == value.category:
            return True
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
        if type(row[1][1]) == str and 'WEEKLY COMMISSION' in row[1][1]:
            if curr_week:
                week_ranges[curr_week] = (curr_start, curr_row-1)
            curr_week = row[1][6]
            curr_start = curr_row
        curr_row += 1
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
            if type(row_data[1]) == str and 'COMPLETED & PAID JOBS' in row_data[1]:
                in_job_section = True
            continue
        # print(type(row_data[1]), row_data[1])
        if type(row_data[1]) == int:
            j_date = row_data[2]
            j_num = row_data[3]
            j_suburb = row_data[4]
            j_subtotal = row_data[5]
            j_materials = row_data[6]
            j_merchantfee = row_data[7]
            j_profit = row_data[8]
            j_paymenttypes = row_data[9].split(', ') if type(row_data[9]) == str else ['N/A']
            j_eftpos = row_data[21]
            j_cash = row_data[22]
            j_paymentplan = row_data[23]
            
            job = JobData(j_num, j_date, j_suburb, j_subtotal, j_materials, j_merchantfee, j_profit, j_paymenttypes, j_eftpos, j_cash, j_paymentplan, curr_cat)
            jobs[j_num] = job

        if type(row_data[1]) == str:
            curr_cat = row_data[1]

    week_data = WeekData(jobs)
    return week_data

def test_job_equality(j1: JobData, j2: JobData):
    """
    Test if jobs are equal, return JobDiff object if not, True if so.
    """
    pass

def main():
    j1 = JobData(1234, date(2025,11,24), 'Mascot', 15564.2, 1235.1, 0, 3000.0, ['Credit Card', 'EFT'], 15564.2*1.1, 0, 0, "COMPLETED & PAID JOBS")
    j2 = JobData(123, date(2025,11,25), 'Mascot', 1234.2, 123.1, 0, 1000.0, ['Cash'], 0, 1234.2*1.1, 0, "COMPLETED & PAID JOBS")
    j3 = JobData(1234, date(2025,11,24), 'Mascot', 15564.2, 1235.1, 0, 3000.0, ['Credit Card', 'EFT'], 15564.2*1.1, 0, 0, "COMPLETED & PAID JOBS")

    # w1 = WeekData({
    #     j1.num: j1,
    #     j2.num: j2
    # })

    print(j1)
    print(j2)
    # print(w1)
    print(j1==j2)
    print(j1==j3)
    # sheet = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=None)
    # week_ranges = get_week_data_ranges(sheet)
    # test_week = list(week_ranges.keys())[1]
    # print(test_week)
    # pprint(extract_jobs_from_week(sheet, week_ranges, test_week))

if __name__ == '__main__':
    main()
