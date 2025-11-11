import datetime as dt
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def is_truthy(v) -> bool:
    return str(v).strip().lower() in ("1", "y", "yes", "true")


def extract_doc_checks(job: dict) -> dict:
    ext = job.get("externalData") or {}
    out = {}
    for k, v in ext.items():
        out[k] = 1 if is_truthy(v) else 0
    return out


def clean_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31] or "Technician"


def _to_date(val):
    if not val:
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, str):
        return dt.datetime.fromisoformat(val.replace("Z", "+00:00")).date()
    if isinstance(val, dt.date):
        return val
    return None


def is_weekend_commissionable(job: dict) -> bool:
    created = _to_date(job.get("createdOn"))
    completed = _to_date(job.get("completedOn") or job.get("closedOn"))
    paid = _to_date(job.get("paidOn") or job.get("invoicePaidOn"))
    sold = created
    if not all([created, completed, paid, sold]):
        return False
    return created == completed == paid == sold


def build_commission_workbook(
    jobs_by_tech: dict[str, list[dict]],
    week_ending: dt.date,
    commission_by_tech: dict[str, str] | None = None,
) -> bytes:
    # discover all doc-check keys
    doc_keys = set()
    for tech_jobs in jobs_by_tech.values():
        for job in tech_jobs:
            ext = job.get("externalData") or {}
            for k in ext:
                doc_keys.add(k)
    doc_keys = sorted(doc_keys)

    wb = Workbook()
    first_sheet_created = False

    for tech_name, jobs in jobs_by_tech.items():
        scheme = commission_by_tech.get(tech_name, "5-or-10") if commission_by_tech else "5-or-10"

        if not first_sheet_created:
            ws = wb.active
            ws.title = clean_sheet_name(tech_name)
            first_sheet_created = True
        else:
            ws = wb.create_sheet(title=clean_sheet_name(tech_name))

        # top summary
        ws["A1"] = "Technician"; ws["B1"] = tech_name
        ws["A2"] = "Week ending"; ws["B2"] = week_ending
        ws["A3"] = "WEEKDAY net profit"
        ws["A4"] = "WEEKDAY commission rate"
        ws["A5"] = "WEEKDAY commission due"
        ws["A6"] = "WEEKEND net profit"
        ws["A7"] = "WEEKEND commission due"
        ws["A8"] = "TOTAL commission"
        ws["A9"] = "Scheme"; ws["B9"] = scheme

        # weekday headers
        weekday_title_row = 11
        weekday_header_row = 12
        ws.cell(row=weekday_title_row, column=1).value = "WEEKDAY JOBS"

        headers = [
            "JOB STATUS",      # 1
            "DATE",            # 2
            "JOB #",           # 3
            "SUBURB",          # 4
            "AMOUNT EXC. GST", # 5
            "MATERIALS",       # 6
            "Merchant Fees",   # 7
            "NET PROFIT",      # 8
            "PAID",            # 9
        ]
        headers.extend(doc_keys)
        headers.append("all_doc_checks_ok")
        headers.append("DESCRIPTION")
        headers.append("NOTES")
        # add weekend-specific helper col at the very end in case row is weekend:
        headers.append("weekend_commissionable")

        for col_idx, h in enumerate(headers, start=1):
            ws.cell(row=weekday_header_row, column=col_idx).value = h

        weekday_row = weekday_header_row + 1
        weekend_header_row = None
        weekend_row = None

        weekday_net_rows = []
        weekend_net_rows = []

        for job in jobs:
            created = _to_date(job.get("createdOn"))
            is_weekend = created is not None and created.weekday() >= 5

            if not is_weekend:
                r = weekday_row
                weekday_row += 1
            else:
                if weekend_header_row is None:
                    weekend_title_row = weekday_row + 1
                    weekend_header_row = weekend_title_row + 1
                    ws.cell(row=weekend_title_row, column=1).value = "WEEKEND JOBS"
                    for col_idx, h in enumerate(headers, start=1):
                        ws.cell(row=weekend_header_row, column=col_idx).value = h
                    weekend_row = weekend_header_row + 1
                r = weekend_row
                weekend_row += 1

            status = job.get("status") or ""
            job_id = job.get("id")
            amount_ex_gst = job.get("subtotal") or job.get("total") or 0
            suburb = None
            loc = job.get("location") or {}
            if isinstance(loc, dict):
                suburb = loc.get("city") or loc.get("name")

            ws.cell(row=r, column=1).value = status
            ws.cell(row=r, column=2).value = created
            ws.cell(row=r, column=3).value = job_id
            ws.cell(row=r, column=4).value = suburb
            ws.cell(row=r, column=5).value = amount_ex_gst

            # net profit formula
            e = get_column_letter(5)
            f = get_column_letter(6)
            g = get_column_letter(7)
            ws.cell(row=r, column=8).value = f"={e}{r}-{f}{r}-{g}{r}"

            # doc checks
            doc_vals = extract_doc_checks(job)
            all_ok = 1
            base_doc_col = 10
            for offset, key in enumerate(doc_keys):
                val = doc_vals.get(key, 0)
                ws.cell(row=r, column=base_doc_col + offset).value = val
                if val != 1:
                    all_ok = 0
            ws.cell(row=r, column=base_doc_col + len(doc_keys)).value = all_ok

            # desc + notes
            ws.cell(row=r, column=base_doc_col + len(doc_keys) + 1).value = job.get("summary") or job.get("name") or ""
            ws.cell(row=r, column=base_doc_col + len(doc_keys) + 2).value = job.get("customerName") or ""

            # weekend helper col (last col)
            weekend_flag_col = base_doc_col + len(doc_keys) + 3
            if is_weekend:
                wc = 1 if is_weekend_commissionable(job) else 0
                ws.cell(row=r, column=weekend_flag_col).value = wc
                weekend_net_rows.append(r)
            else:
                ws.cell(row=r, column=weekend_flag_col).value = 0
                weekday_net_rows.append(r)

        # summaries
        if weekday_net_rows:
            ws["B3"].value = f"=SUM(H{weekday_net_rows[0]}:H{weekday_net_rows[-1]})"
        else:
            ws["B3"].value = 0

        if scheme == "5-or-10":
            ws["B4"].value = "=IF(B3>=25000,0.1,0.05)"
            ws["B5"].value = "=B3*B4"
        else:
            ws["B4"].value = "=IF(B3>=25000,0.1,0)"
            ws["B5"].value = "=B3*B4"

        # weekend: sum only rows with weekend_commissionable = 1
        if weekend_net_rows:
            # weekend_commissionable col index:
            weekend_flag_col_letter = get_column_letter(base_doc_col + len(doc_keys) + 3)
            net_col_letter = "H"
            first_r = weekend_net_rows[0]
            last_r = weekend_net_rows[-1]
            # SUMIFS(Hrange, flagRange, 1)
            ws["B6"].value = (
                f"=SUMIFS({net_col_letter}{first_r}:{net_col_letter}{last_r},"
                f"{weekend_flag_col_letter}{first_r}:{weekend_flag_col_letter}{last_r},1)"
            )
            # 25% of that
            ws["B7"].value = "=B6*0.25"
        else:
            ws["B6"].value = 0
            ws["B7"].value = 0

        # total commission
        ws["B8"].value = "=B5+B7"

        # tidy widths
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 16

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()
