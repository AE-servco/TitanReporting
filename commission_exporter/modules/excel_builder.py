import datetime as dt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.formatting.rule import CellIsRule

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


    # =========================== STYLING ===========================

    thin_border = Side(style='thin', color='000000') # Black thin border
    med_border = Side(style='medium', color='000000') # Black medium border
    thick_border = Side(style='thick', color='000000') # Black thick border
    double_border = Side(style='double', color='000000')

    yellow_fill = PatternFill(patternType='solid', fgColor='FFFF00')

    font_red = Font(color='FF0000')

    cell_border_full = Border(top=thin_border, bottom=thin_border, right=thin_border, left=thin_border)
    border_double_bottom = Border(bottom=double_border)

    cell_border = {
        'topleft': Border(left=med_border, top=med_border),
        'topright': Border(right=med_border, top=med_border),
        'bottomleft': Border(left=med_border, bottom=med_border),
        'bottomright': Border(right=med_border, bottom=med_border),
        'bottomtop': Border(top=med_border, bottom=med_border),
        'bottomtopright': Border(top=med_border, bottom=med_border, right=med_border),
        'bottom': Border(bottom=med_border),
        'top': Border(top=med_border),
    }

    align_center = Alignment(horizontal='center')

    # =========================== WORKBOOK ===========================

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
        ws.cell(2,22, 'EFTPOS').border = cell_border['bottomtop']
        ws.cell(2,23, 'CASH').border = cell_border['bottomtop']
        ws.cell(2,24, 'Payment Plan').border = cell_border['bottomtopright']
        # =========================================

        curr_row = 3

        for cat, cat_text in CATEGORY_ORDER.items():
            ws.cell(curr_row,1,cat_text)
            curr_row += 1
            cat_row_start = curr_row

            if cat == 'prev':
                for row in ws[f'A{curr_row-1}:X{curr_row+7}']:
                    for cell in row:
                        cell.fill = yellow_fill
                for i in range(1,9):
                    ws.cell(curr_row, 1, i)
                    curr_row += 1

            else:
                jobs = job_cats.get(cat, [])
                if not jobs:
                    curr_row += 1
                job_count = 1
                for job in jobs:
                    ws.cell(curr_row, 1, job_count)
                    ws.cell(curr_row, 2, job['created_str'])
                    ws.cell(curr_row, 3, job['num'])
                    ws.cell(curr_row, 4, job['suburb'])
                    ws.cell(curr_row, 5, job['subtotal']).number_format = '$ 00.00'
                    if not job['unsuccessful']:
                        ws.cell(curr_row, 6, job['subtotal']*0.2).number_format = '$ 00.00'
                    # 7
                    ws.cell(curr_row, 8, f"={get_column_letter(5)}{curr_row}-{get_column_letter(6)}{curr_row}-{get_column_letter(7)}{curr_row}").number_format = '$ 00.00'
                    ws.cell(curr_row, 9, job['payment_types'])
                    # 10 TODO: all doc checks complete
                    doc_check_complete_col = f'=IF(OR({"=0, ".join([f"{get_column_letter(i)}{curr_row}" for i in range(11,20)])}), "N","Y")'
                    ws.cell(curr_row, 10, doc_check_complete_col).alignment = align_center
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

                    
                    
                    payments = job.get('payment_amt')

                    if type(payments) == str:
                        payments = payments.split(', ')
                        
                        # All payment types:
                        # {
                            # 'AMEX',
                            # 'Applied Payment for AR',
                            # 'Cash',
                            # 'Check',
                            # 'Credit Card',
                            # 'EFT/Bank Transfer',
                            # 'Humm - Finance Fee',
                            # 'Humm Payment Plan',
                            # 'Imported Default Credit Card',
                            # 'MasterCard',
                            # 'Payment Plan',
                            # 'Payment Plan - Fee',
                            # 'Processed in ServiceM8',
                            # 'Refund (check)',
                            # 'Refund (credit card)',
                            # 'Visa'
                        # }
                        
                        p_types = {
                            'Cr': [],
                            'Ca': [],
                            'EF': [],
                        }
                        for p in payments:
                            try:
                                p_types[p[:2]].append(p[2:])
                            except:
                                continue
                        if p_types['EF'] or p_types['Cr']:
                            ws.cell(curr_row, 22, f"={'+'.join(p_types['EF'] + p_types['Cr'])}")
                        if p_types['Ca']:
                            ws.cell(curr_row, 23, f"={'+'.join(p_types['Ca'])}")
                    curr_row += 1
                    job_count += 1
            
            # Doc check formatting 0s to red
            dc_start_col_letter = get_column_letter(10)
            dc_end_col_letter = get_column_letter(20)
            # dc_range = "A1:A100"
            dc_range = f'{dc_start_col_letter}{cat_row_start}:{dc_end_col_letter}{curr_row-1}'
            ws.conditional_formatting.add(
                dc_range, 
                CellIsRule(operator="equal", formula=["0"], font=font_red)
            )

            # totals row
            amt_col = 5
            amt_letter = get_column_letter(amt_col)
            materials_col = 6
            materials_letter = get_column_letter(materials_col)
            merchantf_col = 7
            merchantf_letter = get_column_letter(merchantf_col)
            profit_col = 8
            profit_letter = get_column_letter(profit_col)
            ws.cell(curr_row, amt_col, f"=SUM({amt_letter}{cat_row_start}:{amt_letter}{curr_row-1})").number_format = '$ 00.00'
            ws.cell(curr_row, materials_col, f"=SUM({materials_letter}{cat_row_start}:{materials_letter}{curr_row-1})").number_format = '$ 00.00'
            ws.cell(curr_row, merchantf_col, f"=SUM({merchantf_letter}{cat_row_start}:{merchantf_letter}{curr_row-1})").number_format = '$ 00.00'
            ws.cell(curr_row, profit_col, f"=SUM({profit_letter}{cat_row_start}:{profit_letter}{curr_row-1})").number_format = '$ 00.00'
            ws.cell(curr_row, amt_col).border = border_double_bottom
            ws.cell(curr_row, materials_col).border = border_double_bottom
            ws.cell(curr_row, merchantf_col).border = border_double_bottom
            ws.cell(curr_row, profit_col).border = border_double_bottom
            curr_row +=2
        
        for row in ws[f'V{3}:X{curr_row-3}']:
            for cell in row:
                cell.border = cell_border_full

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()