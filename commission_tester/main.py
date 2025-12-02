from dataclasses import dataclass, field
import pandas as pd
from datetime import date, time, datetime

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
    num: int
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
    pass

def extract_summary_from_week(week_end_date: date):
    """
    Gets summary data from the week specified.
    """
    pass

def extract_jobs_from_week(week_end_date: date):
    """
    Gets job data from the week specified.
    """
    pass

def test_job_equality(j1: JobData, j2: JobData):
    """
    Test if jobs are equal, return JobDiff object if not, True if so.
    """
    pass

def main():
    j1 = JobData(1234, date(2025,11,24), 'Mascot', 15564.2, 1235.1, 0, 3000.0, ['Credit Card', 'EFT'], 15564.2*1.1, 0, 0, "COMPLETED & PAID JOBS")
    j2 = JobData(123, date(2025,11,25), 'Mascot', 1234.2, 123.1, 0, 1000.0, ['Cash'], 0, 1234.2*1.1, 0, "COMPLETED & PAID JOBS")
    j3 = JobData(1234, date(2025,11,24), 'Mascot', 15564.2, 1235.1, 0, 3000.0, ['Credit Card', 'EFT'], 15564.2*1.1, 0, 0, "COMPLETED & PAID JOBS")

    w1 = WeekData({
        j1.num: j1,
        j2.num: j2
    })

    print(j1)
    print(j2)
    print(w1)
    print(j1==j2)
    print(j1==j3)

if __name__ == '__main__':
    main()
