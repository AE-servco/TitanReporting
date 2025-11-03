import holidays
import datetime

aus_holidays = holidays.Australia()
nsw_holidays = holidays.Australia(subdiv='NSW')

date = datetime.date(2025, 10, 6)
if date in aus_holidays:
    print(f"{date} is {aus_holidays[date]}")
if date in nsw_holidays:
    print(f"{date} is {nsw_holidays[date]}")

combined = holidays.HolidayBase()
combined.update(aus_holidays)
combined.update(nsw_holidays)

if date in combined:
    print(f"{date} is {combined[date]}")