import datetime as dt
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side

def build_workbook(
    jobs_by_tech: dict[str, list[dict]],
    week_ending: dt.date,
) -> bytes:
    
    CATEGORY_ORDER = {
        'wk_complete_paid': 'COMPLETED & PAID JOBS', 
        'wkend_complete_paid': 'WEEKEND COMPLETED & PAID JOBS', 
        'prev': 'PREVIOUS JOBS COMPLETED & PAID (COMMISSION) - Modoras Team Please do ADD to THIS SECTION or AMEND/TOUCH', 
        'wk_complete_unpaid': 'CURRENT JOBS COMPLETED (AWAITING PAYMENT)', 
        'wkend_complete_unpaid': 'WEEKEND CURRENT JOBS COMPLETED (AWAITING PAYMENT)', 
        'wk_wo': 'CURRENT WORK ORDERS (AWAITING PAYMENT)', 
        'wkend_wo': 'WEEKEND WORK ORDERS (AWAITING PAYMENT)', 
        'wk_unsucessful': 'UNSUCCESSFUL JOBS',
        'wkend_unsucessful': 'WEEKEND UNSUCCESSFUL JOBS'
    }

    thin_border = Side(style='thin', color='000000') # Black thin border
    med_border = Side(style='medium', color='000000') # Black medium border
    thick_border = Side(style='thick', color='000000') # Black thick border

    cell_border = {
        'topleft': Border(left=med_border, top=med_border),
        'topright': Border(right=med_border, top=med_border),
        'bottomleft': Border(left=med_border, bottom=med_border),
        'bottomright': Border(right=med_border, bottom=med_border),
        'bottom': Border(bottom=med_border),
        'top': Border(top=med_border),
    }

    wb = Workbook()
    
    first_sheet_created = False

    for tech, job_cats in jobs_by_tech.items():

        if not first_sheet_created:
            ws = wb.active
            ws.title = tech
            first_sheet_created = True
        else:
            ws = wb.create_sheet(title=tech)

        # ================ HEADERS ==================
        ws.cell(1,1, 'JOB DETAILS').border = cell_border['topleft']
        ws.cell(1,2).border = cell_border['top']
        ws.cell(1,3).border = cell_border['top']
        ws.cell(1,4).border = cell_border['top']
        ws.cell(1,5, 'JOB AMOUNT').border = cell_border['topleft']
        ws.cell(1,6).border = cell_border['top']
        ws.cell(1,7).border = cell_border['top']
        ws.cell(1,8).border = cell_border['top']
        ws.cell(1,9).border = cell_border['top']
        ws.cell(1,11, 'PHOTOS').border = cell_border['topleft']
        ws.cell(1,12).border = cell_border['top']
        ws.cell(1,13).border = cell_border['top']
        ws.cell(1,14, 'QUOTE').border = cell_border['topleft']
        ws.cell(1,15).border = cell_border['top']
        ws.cell(1,16).border = cell_border['top']
        ws.cell(1,17, 'INVOICE').border = cell_border['topleft']
        ws.cell(1,18).border = cell_border['top']
        ws.cell(1,19).border = cell_border['top']
        ws.cell(1,20).border = cell_border['top']
        ws.cell(1,21, 'NOTES').border = cell_border['topright']

        ws.cell(2,1, 'JOB STATUS').border = cell_border['bottomleft']
        ws.cell(2,2, 'DATE').border = cell_border['bottom']
        ws.cell(2,3, 'JOB #').border = cell_border['bottom']
        ws.cell(2,4, 'SUBURB').border = cell_border['bottomright']
        ws.cell(2,5, 'AMOUNT EXC. GST').border = cell_border['bottom']
        ws.cell(2,6, 'MATERIALS').border = cell_border['bottom']
        ws.cell(2,7, 'MERCHANT FEES').border = cell_border['bottom']
        ws.cell(2,8, 'NET PROFIT').border = cell_border['bottom']
        ws.cell(2,9, 'PAID').border = cell_border['bottomright']
        ws.cell(2,10, 'DOC CHECK DONE').border = cell_border['bottomright']
        ws.cell(2,11, 'BEFORE').border = cell_border['bottom']
        ws.cell(2,12, 'AFTER').border = cell_border['bottom']
        ws.cell(2,13, 'RECEIPT').border = cell_border['bottomright']
        ws.cell(2,14, 'DESCRIPTION').border = cell_border['bottom']
        ws.cell(2,15, 'SIGNED').border = cell_border['bottom']
        ws.cell(2,16, 'EMAILED').border = cell_border['bottomright']
        ws.cell(2,17, 'DESCRIPTION').border = cell_border['bottom']
        ws.cell(2,18, 'SIGNED').border = cell_border['bottom']
        ws.cell(2,19, 'EMAILED').border = cell_border['bottomright']
        ws.cell(2,20, '5 Star Review').border = cell_border['bottom']
        ws.cell(2,21).border = cell_border['bottomright']
        # =========================================

        curr_row = 3

        for cat, cat_text in CATEGORY_ORDER.items():
            ws.cell(curr_row,1,cat_text)
            curr_row += 1
            if cat == 'prev':
                for i in range(1,9):
                    ws.cell(curr_row, 1, i)
            else:
                jobs = job_cats.get(cat, [])
                job_count = 1
                for job in jobs:
                    ws.cell(curr_row, 1, job_count)
                    ws.cell(curr_row, 2, job['created_str'])
                    ws.cell(curr_row, 3, job['num'])
                    ws.cell(curr_row, 4, job['suburb'])
                    ws.cell(curr_row, 5, job['subtotal'])
                    ws.cell(curr_row, 6, job['subtotal']*0.2)
                    # 7
                    # 8
                    ws.cell(curr_row, 9, job['payment_types'])
                    # 10 TODO: all doc checks complete
                    ws.cell(curr_row, 11, job['Before Photo'])
                    ws.cell(curr_row, 12, job['After Photo'])
                    ws.cell(curr_row, 13, job['Receipt Photo'])
                    ws.cell(curr_row, 14, job['Quote Description'])
                    ws.cell(curr_row, 15, job['Quote Description'])
                    ws.cell(curr_row, 16, job['Quote Emailed'])
                    ws.cell(curr_row, 17, job['Invoice Description'])
                    ws.cell(curr_row, 18, job['Invoice Signed'])
                    ws.cell(curr_row, 19, job['Invoice Emailed'])
                    ws.cell(curr_row, 20, job['5 Star Review'])

                    curr_row += 1
                    job_count += 1
            
            # TODO: Add totals row
            curr_row +=2


    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()