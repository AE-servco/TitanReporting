import datetime as dt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule

def formatted_cell(worksheet, row: int, col: int, val: str | None=None, font: Font | None=None, border: Border | None=None, number_format: str | None=None, fill: PatternFill | None=None):
    if val:
        cell = worksheet.cell(row, col, val)
    else:
        cell = worksheet.cell(row, col)
    if font:
        cell.font = font
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = fill

    return cell

def build_workbook(
    jobs_by_tech: dict[str, list[dict]],
    end_date: dt.date
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
        'wkend_unsucessful': 'WEEKEND UNSUCCESSFUL JOBS',
        'wk_uncategorised': 'WEEK UNCATEGORISED',
        'wkend_uncategorised': 'WEEKEND UNCATEGORISED',
    }
    dates = {
        'monday': end_date - dt.timedelta(days=6),
        'tuesday': end_date - dt.timedelta(days=5),
        'wednesday': end_date - dt.timedelta(days=4),
        'thursday': end_date - dt.timedelta(days=3),
        'friday': end_date - dt.timedelta(days=2),
        'saturday': end_date - dt.timedelta(days=1),
    }
    
    date_strs = {day: date.strftime("%d/%m/%Y") for day, date in dates.items()}
    # monday_str = monday.strftime("%d/%m/%Y")
    # tuesday_str = tuesday.strftime("%d/%m/%Y")
    # wednesday_str = wednesday.strftime("%d/%m/%Y")
    # thursday_str = thursday.strftime("%d/%m/%Y")
    # friday_str = friday.strftime("%d/%m/%Y")
    # saturday_str = saturday.strftime("%d/%m/%Y")

    cats_count_for_total = [
        'wk_complete_paid',
        'wkend_complete_paid',
        'wk_complete_unpaid',
        'wkend_complete_unpaid',
        'wk_wo',
        'wkend_wo',
    ]

    cats_count_awaiting_pay_wk = [
        'wk_complete_unpaid',
        'wk_wo',
    ]

    cats_count_awaiting_pay_wkend = [
        'wkend_complete_unpaid',
        'wkend_wo',
    ]

    cats_count_for_potential_wk = [
        'wk_complete_paid',
        'wk_complete_unpaid',
        'wk_wo',
        'wk_unsucessful',
    ]

    cats_count_for_potential_wkend = [
        'wkend_complete_paid',
        'wkend_complete_unpaid',
        'wkend_wo',
        'wkend_unsucessful',
    ]

    cats_count_for_unsuccessful = [
        'wk_unsucessful',
        'wkend_unsucessful',
    ]

    # =========================== STYLING ===========================

    thin_border = Side(style='thin', color='000000') # Black thin border
    med_border = Side(style='medium', color='000000') # Black medium border
    thick_border = Side(style='thick', color='000000') # Black thick border
    double_border = Side(style='double', color='000000')

    yellow_fill = PatternFill(patternType='solid', fgColor='FFFF00')
    blue_fill = PatternFill(patternType='solid', fgColor='3399F2')

    font_red = Font(color='FF0000')
    font_red_bold = Font(color='FF0000', bold=True)
    font_green_bold = Font(color='1F9C40', bold=True)
    font_green = Font(color='1F9C40')
    font_bold = Font(bold=True)

    cell_border_full = Border(top=thin_border, bottom=thin_border, right=thin_border, left=thin_border)
    border_double_bottom = Border(bottom=double_border)

    cell_border = {
        'topleft': Border(left=med_border, top=med_border),
        'topright': Border(right=med_border, top=med_border),
        'topleftright': Border(left=med_border, top=med_border, right=med_border),
        'bottomleft': Border(left=med_border, bottom=med_border),
        'bottomright': Border(right=med_border, bottom=med_border),
        'bottomtop': Border(top=med_border, bottom=med_border),
        'bottomtopright': Border(top=med_border, bottom=med_border, right=med_border),
        'bottomleftright': Border(left=med_border, bottom=med_border, right=med_border),
        'bottom': Border(bottom=med_border),
        'top': Border(top=med_border),
        'right': Border(right=med_border),
        'left': Border(left=med_border),
        'leftright': Border(left=med_border, right=med_border),
    }

    align_center = Alignment(horizontal='center')

    # =========================== WORKBOOK ===========================

    wb = Workbook()
    
    first_sheet_created = False

    for tech in sorted(jobs_by_tech.keys()):
        job_cats = jobs_by_tech[tech]
        if not first_sheet_created:
            ws = wb.active
            ws.title = tech[:-1]
            first_sheet_created = True
        else:
            ws = wb.create_sheet(title=tech[:-1])

        sales_color = '3da813'
        installer_color = 'e3b019'
        unknown_color = '8093f2'

        if tech[-1] == 'S':
            ws.sheet_properties.tabColor = sales_color
        elif tech[-1] == 'I':
            ws.sheet_properties.tabColor = installer_color
        else:
            ws.sheet_properties.tabColor = unknown_color

        summary_top_row = 1

        # ================ SUMMARY ==================
        formatted_cell(ws, summary_top_row,1, 'WEEKLY COMMISSION', font = font_bold)
        formatted_cell(ws, summary_top_row,5, 'WEEK ENDING', font = font_red_bold)
        formatted_cell(ws, summary_top_row,6, 'PUT DATE HERE', font = font_red_bold)

        # box 1
        formatted_cell(ws, summary_top_row + 2,1, 'BOOKED JOBS',font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 2,2, '=B4+B5', font = font_bold, border = cell_border['topright'])
        formatted_cell(ws, summary_top_row + 3,1, 'SUCCESSFUL', font = font_bold)
        formatted_cell(ws, summary_top_row + 3,2, '=O11', font = font_bold, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 4,1, 'UNSUCCESSFUL', font = font_bold)
        formatted_cell(ws, summary_top_row + 4,2, '=O12', font = font_bold, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 5,1, 'SUCESSFUL (%)', font = font_bold)
        formatted_cell(ws, summary_top_row + 5,2, '=O13', font = font_bold, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 6,2, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 7,1, 'AVERAGE SALE', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 7,2, '=O14', font = font_bold, border = cell_border['bottomright'])
        
        # box 2
        formatted_cell(ws, summary_top_row + 2,4, 'WEEKLY TARGET', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 2,5, border = cell_border['top'])        
        formatted_cell(ws, summary_top_row + 2,6, border = cell_border['topright'])        
        formatted_cell(ws, summary_top_row + 3,4, 'Tier 1', border = cell_border['left'])        
        formatted_cell(ws, summary_top_row + 3,5, '<$25000')        
        formatted_cell(ws, summary_top_row + 3,6, '5%', border = cell_border['right'])        
        formatted_cell(ws, summary_top_row + 4,4, 'Tier 2', border = cell_border['bottomleft'])        
        formatted_cell(ws, summary_top_row + 4,5, '>=$25000', border = cell_border['bottom'])        
        formatted_cell(ws, summary_top_row + 4,6, '10%', border = cell_border['bottomright'])    

        # box 3
        formatted_cell(ws, summary_top_row + 6,5, 'ACTUAL', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 6,6, 'POTENTIAL', font = font_bold, border = cell_border['topright'])        
        formatted_cell(ws, summary_top_row + 6,7, 'Exc. SUPER', font = font_green_bold)        
        formatted_cell(ws, summary_top_row + 7,4, 'NET PROFIT', font = font_bold, border = cell_border['topleft'])        
        formatted_cell(ws, summary_top_row + 7,5, '=O10', border = cell_border['topleft'])        
        formatted_cell(ws, summary_top_row + 8,4, 'UNLOCKED', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 8,5, '=IF(E8>=25000,10%,5%)', border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 9,4, 'COMMISSION - PAY OUT', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 9,5, '=E8*E9', border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 9,7, '=E10/1.12', font = font_green_bold)
        formatted_cell(ws, summary_top_row + 10,4, 'EMERGENCY', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 10,5, '=P10', border = cell_border['left'])        
        formatted_cell(ws, summary_top_row + 11,4, 'EMERGENCY - PAY OUT', font = font_bold, border = cell_border['left'])        
        formatted_cell(ws, summary_top_row + 11,5, '=E11*0.25', border = cell_border['left'])        
        formatted_cell(ws, summary_top_row + 11,7, '=E12/1.12', font = font_green_bold)
        formatted_cell(ws, summary_top_row + 12,4, 'PREV. WEEK', font = font_bold, border = cell_border['left'])        
        formatted_cell(ws, summary_top_row + 12,5, 0, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 13,4, 'PREV. WEEK - PAY OUT', font = font_bold, border = cell_border['bottomleft'])        
        formatted_cell(ws, summary_top_row + 13,5, '=E13*0.05', border = cell_border['bottomleft'])        
        formatted_cell(ws, summary_top_row + 13,7, '=E14/1.12', font = font_green_bold)
        formatted_cell(ws, summary_top_row + 14,4, '5 Star Review', font = font_green_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 14,5, '=E16*50', border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 15,4, '5 Star Notes', font = font_green_bold, border = cell_border['bottomleft'])        
        formatted_cell(ws, summary_top_row + 15,5, '=S4', border = cell_border['left'])

        formatted_cell(ws, summary_top_row + 8,6, '==IF(F8>=25000,10%,5%)', font=font_red, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 9,6, '=F8*F9', font = font_red, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 11,6, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 12,6, 0, font = font_red, border = cell_border['right'])
        formatted_cell(ws, summary_top_row + 13,6, '==F13*0.05', font = font_red, border = cell_border['bottomright'])
        
        #box 4
        formatted_cell(ws, summary_top_row + 1,10, 'PHOTOS', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 1,11, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 1,12, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 1,13, 'QUOTE', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 1,14, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 1,15, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 1,16, 'INVOICE', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 1,17, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 1,18, border = cell_border['topright'])
        formatted_cell(ws, summary_top_row + 1,19, border = cell_border['topright'])
        formatted_cell(ws, summary_top_row + 2,10, 'BEFORE', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 2,11, 'AFTER', font = font_bold, border = cell_border['bottom'])
        formatted_cell(ws, summary_top_row + 2,12, 'RECEIPT', font = font_bold, border = cell_border['bottom'])
        formatted_cell(ws, summary_top_row + 2,13, 'DESCRIPTION', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 2,14, 'SIGNED', font = font_bold, border = cell_border['bottom'])
        formatted_cell(ws, summary_top_row + 2,15, 'EMAILED', font = font_bold, border = cell_border['bottom'])
        formatted_cell(ws, summary_top_row + 2,16, 'DESCRIPTION', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 2,17, 'SIGNED', font = font_bold, border = cell_border['bottom'])
        formatted_cell(ws, summary_top_row + 2,18, 'EMAILED', font = font_bold, border = cell_border['bottom'])
        formatted_cell(ws, summary_top_row + 2,19, '5 Star Review', font = font_bold, border = cell_border['bottomleftright'])
        formatted_cell(ws, summary_top_row + 3,9, 'TAKEN', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 4,9, '%', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,19, border = cell_border['bottomleftright'])

        # box 5
        formatted_cell(ws, summary_top_row + 7,10, 'DAILY NET PROFIT', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 7,11, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 7,12, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 7,13, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 7,14, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 7,15, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 7,16, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 7,17, 'WEEKDAY', font = font_red_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 7,18, 'WEEKEND', font = font_red_bold, border = cell_border['topleftright'])
        formatted_cell(ws, summary_top_row + 8,10, 'MON', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 8,11, 'TUE', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 8,12, 'WED', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 8,13, 'THU', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 8,14, 'FRI', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 8,15, 'Total', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 8,16, 'SAT', font = font_green_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 8,17, 'Awaiting Payment', font = font_red_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 8,18, 'Awaiting Payment', font = font_red_bold, border = cell_border['topleftright'])
        ##
        formatted_cell(ws, summary_top_row + 9,15, '=SUM(J10:N10)', font = font_bold, border = cell_border['bottom'])
        ##

        formatted_cell(ws, summary_top_row + 10,9, 'Successful')
        formatted_cell(ws, summary_top_row + 11,9, 'Unsuccessful')
        formatted_cell(ws, summary_top_row + 12,9, 'Success rate')
        formatted_cell(ws, summary_top_row + 13,9, 'Avg sale')

        formatted_cell(ws, summary_top_row + 12,10, '=J11/(J11+J12)')
        formatted_cell(ws, summary_top_row + 13,10, '=J10/J11')

        formatted_cell(ws, summary_top_row + 12,11, '=K11/(K11+K12)')
        formatted_cell(ws, summary_top_row + 13,11, '=K10/K11')
        
        formatted_cell(ws, summary_top_row + 12,12, '=L11/(L11+L12)')
        formatted_cell(ws, summary_top_row + 13,12, '=L10/L11')
        
        formatted_cell(ws, summary_top_row + 12,13, '=M11/(M11+M12)')
        formatted_cell(ws, summary_top_row + 13,13, '=M10/M11')
        
        formatted_cell(ws, summary_top_row + 12,14, '=N11/(N11+N12)')
        formatted_cell(ws, summary_top_row + 13,14, '=N10/N11')
        
        formatted_cell(ws, summary_top_row + 10,15, '=SUM(J11:N11)') 
        formatted_cell(ws, summary_top_row + 11,15, '=SUM(J12:N12)')
        formatted_cell(ws, summary_top_row + 12,15, '=O11/(O11+O12)')
        formatted_cell(ws, summary_top_row + 13,15, '=O10/O11')
        
        formatted_cell(ws, summary_top_row + 12,16, '=P11/(P11+P12)')
        formatted_cell(ws, summary_top_row + 13,16, '=P10/P11')

        # box 6 - management summary
        formatted_cell(ws, summary_top_row + 15,10, 'Management Summary - SALES ONLY - COLUMN F', font = font_bold, border = cell_border['topleft'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,11, border = cell_border['top'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,12, border = cell_border['top'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,13, border = cell_border['top'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,14, border = cell_border['top'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,15, border = cell_border['top'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,16, border = cell_border['top'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,17, 'WEEKDAY', font = font_red_bold, border = cell_border['topleft'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 15,18, 'WEEKEND', font = font_red_bold, border = cell_border['topleftright'], fill=blue_fill)
        formatted_cell(ws, summary_top_row + 16,10, 'MON', font = font_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 16,11, 'TUE', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 16,12, 'WED', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 16,13, 'THU', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 16,14, 'FRI', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 16,15, 'Total', font = font_bold, border = cell_border['top'])
        formatted_cell(ws, summary_top_row + 16,16, 'SAT', font = font_green_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 16,17, 'Awaiting Payment', font = font_red_bold, border = cell_border['topleft'])
        formatted_cell(ws, summary_top_row + 16,18, 'Awaiting Payment', font = font_red_bold, border = cell_border['topleftright'])
        ##
        formatted_cell(ws, summary_top_row + 17,15, '=SUM(J18:N18)', font = font_bold, border = cell_border['bottom'])#
        ##

        formatted_cell(ws, summary_top_row + 18,9, 'Successful')
        formatted_cell(ws, summary_top_row + 19,9, 'Unsuccessful')
        formatted_cell(ws, summary_top_row + 20,9, 'Success rate')
        formatted_cell(ws, summary_top_row + 21,9, 'Avg sale')

        formatted_cell(ws, summary_top_row + 18,10, '=J11') 
        formatted_cell(ws, summary_top_row + 19,10, '=J12')
        formatted_cell(ws, summary_top_row + 20,10, '=J13')
        formatted_cell(ws, summary_top_row + 21,10, '=J14')

        formatted_cell(ws, summary_top_row + 18,11, '=K11') 
        formatted_cell(ws, summary_top_row + 19,11, '=K12')
        formatted_cell(ws, summary_top_row + 20,11, '=K13')
        formatted_cell(ws, summary_top_row + 21,11, '=K14')
        
        formatted_cell(ws, summary_top_row + 18,12, '=L11') 
        formatted_cell(ws, summary_top_row + 19,12, '=L12')
        formatted_cell(ws, summary_top_row + 20,12, '=L13')
        formatted_cell(ws, summary_top_row + 21,12, '=L14')
        
        formatted_cell(ws, summary_top_row + 18,13, '=M11') 
        formatted_cell(ws, summary_top_row + 19,13, '=M12')
        formatted_cell(ws, summary_top_row + 20,13, '=M13')
        formatted_cell(ws, summary_top_row + 21,13, '=M14')
        
        formatted_cell(ws, summary_top_row + 18,14, '=N11') 
        formatted_cell(ws, summary_top_row + 19,14, '=N12')
        formatted_cell(ws, summary_top_row + 20,14, '=N13')
        formatted_cell(ws, summary_top_row + 21,14, '=N14')
        
        formatted_cell(ws, summary_top_row + 18,15, '=O11') 
        formatted_cell(ws, summary_top_row + 19,15, '=O12')
        formatted_cell(ws, summary_top_row + 20,15, '=O13')
        formatted_cell(ws, summary_top_row + 21,15, '=O14')
        
        formatted_cell(ws, summary_top_row + 18,16, '=P11') 
        formatted_cell(ws, summary_top_row + 19,16, '=P12')
        formatted_cell(ws, summary_top_row + 20,16, '=P13')
        formatted_cell(ws, summary_top_row + 21,16, '=P14')

        # curr_row += 1

        # ws.cell(curr_row,10, 'PHOTOS').font = font_bold
        # ws.cell(curr_row,10).border = cell_border['topleft']
        # ws.cell(curr_row,11).border = cell_border['top']
        # ws.cell(curr_row,12).border = cell_border['top']
        # ws.cell(curr_row,13, 'QUOTE').font = font_bold
        # ws.cell(curr_row,13).border = cell_border['topleft']
        # ws.cell(curr_row,14).border = cell_border['top']
        # ws.cell(curr_row,15).border = cell_border['top']
        # ws.cell(curr_row,16, 'INVOICE').font = font_bold
        # ws.cell(curr_row,16).border = cell_border['topleft']
        # ws.cell(curr_row,17).border = cell_border['top']
        # ws.cell(curr_row,18).border = cell_border['topright']
        # ws.cell(curr_row,19).border = cell_border['topright']

        # curr_row += 1

        # ws.cell(curr_row,4, 'WEEKLY TARGET').font = font_bold
        # ws.cell(curr_row,4).border = cell_border['topleft']
        # ws.cell(curr_row,5).border = cell_border['top']
        # ws.cell(curr_row,6).border = cell_border['topright']
        # ws.cell(curr_row,10, 'BEFORE').font = font_bold

        curr_row = 26
        # ================ HEADERS ==================
        ws.cell(curr_row,1, 'JOB DETAILS').border = cell_border['topleft']
        ws.cell(curr_row,2).border = cell_border['top']
        ws.cell(curr_row,3).border = cell_border['top']
        ws.cell(curr_row,4).border = cell_border['top']
        ws.cell(curr_row,5, 'JOB AMOUNT').border = cell_border['topleft']
        ws.cell(curr_row,6).border = cell_border['top']
        ws.cell(curr_row,7).border = cell_border['top']
        ws.cell(curr_row,8).border = cell_border['top']
        ws.cell(curr_row,9).border = cell_border['top']
        ws.cell(curr_row,10).border = cell_border['topleft']
        ws.cell(curr_row,11, 'PHOTOS').border = cell_border['topleft']
        ws.cell(curr_row,12).border = cell_border['top']
        ws.cell(curr_row,13).border = cell_border['top']
        ws.cell(curr_row,14, 'QUOTE').border = cell_border['topleft']
        ws.cell(curr_row,15).border = cell_border['top']
        ws.cell(curr_row,16).border = cell_border['top']
        ws.cell(curr_row,17, 'INVOICE').border = cell_border['topleft']
        ws.cell(curr_row,18).border = cell_border['top']
        ws.cell(curr_row,19).border = cell_border['top']
        ws.cell(curr_row,20).border = cell_border['top']
        ws.cell(curr_row,21, 'NOTES').border = cell_border['topright']

        curr_row += 1

        ws.cell(curr_row,1, 'JOB STATUS').border = cell_border['bottomleft']
        ws.cell(curr_row,2, 'DATE').border = cell_border['bottom']
        ws.cell(curr_row,3, 'JOB #').border = cell_border['bottom']
        ws.cell(curr_row,4, 'SUBURB').border = cell_border['bottomright']
        ws.cell(curr_row,5, 'AMOUNT EXC. GST').border = cell_border['bottom']
        ws.cell(curr_row,6, 'MATERIALS').border = cell_border['bottom']
        ws.cell(curr_row,7, 'MERCHANT FEES').border = cell_border['bottom']
        ws.cell(curr_row,8, 'NET PROFIT').border = cell_border['bottom']
        ws.cell(curr_row,9, 'PAID').border = cell_border['bottomright']
        ws.cell(curr_row,10, 'DOC CHECK DONE').border = cell_border['bottomright']
        ws.cell(curr_row,11, 'BEFORE').border = cell_border['bottom']
        ws.cell(curr_row,12, 'AFTER').border = cell_border['bottom']
        ws.cell(curr_row,13, 'RECEIPT').border = cell_border['bottomright']
        ws.cell(curr_row,14, 'DESCRIPTION').border = cell_border['bottom']
        ws.cell(curr_row,15, 'SIGNED').border = cell_border['bottom']
        ws.cell(curr_row,16, 'EMAILED').border = cell_border['bottomright']
        ws.cell(curr_row,17, 'DESCRIPTION').border = cell_border['bottom']
        ws.cell(curr_row,18, 'SIGNED').border = cell_border['bottom']
        ws.cell(curr_row,19, 'EMAILED').border = cell_border['bottomright']
        ws.cell(curr_row,20, '5 Star Review').border = cell_border['bottom']
        ws.cell(curr_row,21).border = cell_border['bottomright']
        ws.cell(curr_row,22, 'EFTPOS').border = cell_border['bottomtop']
        ws.cell(curr_row,23, 'CASH').border = cell_border['bottomtop']
        ws.cell(curr_row,24, 'Payment Plan').border = cell_border['bottomtopright']
        
        curr_row += 1
        # =========================================

        
        # Summary input variables 
        # =========================================

        job_counts_per_day = {
            '0': 0,
            '1': 0,
            '2': 0,
            '3': 0,
            '4': 0,
            '5': 0,
            '6': 0,
            '7': 0,
        }

        sales_per_day = {
            '0': 0,
            '1': 0,
            '2': 0,
            '3': 0,
            '4': 0,
            '5': 0,
            '6': 0,
            '7': 0,
        }

        profit_per_day = {
            '0': 0,
            '1': 0,
            '2': 0,
            '3': 0,
            '4': 0,
            '5': 0,
            '6': 0,
            '7': 0,
        }
        monday_jobs = 0
        tuesday_jobs = 0
        wednesday_jobs = 0
        thursday_jobs = 0
        friday_jobs = 0
        saturday_jobs = 0
        
        monday_sales = 0
        tuesday_sales = 0
        wednesday_sales = 0
        thursday_sales = 0
        friday_sales = 0
        saturday_sales = 0
        
        monday_profit = 0
        tuesday_profit = 0
        wednesday_profit = 0
        thursday_profit = 0
        friday_profit = 0
        saturday_profit = 0

        cat_row_info = {}

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
                cat_font = None
                if cat in ['wk_complete_unpaid', 'wk_wo']:
                    cat_font = font_green
                for job in jobs:
                    # if cat in cats_count_for_total:

                    ws.cell(curr_row, 1, job_count)
                    ws.cell(curr_row, 2, job['created_str'])
                    ws.cell(curr_row, 3, job['num'])
                    ws.cell(curr_row, 4, job['suburb'])
                    if not job['unsuccessful']:
                        ws.cell(curr_row, 5, job['inv_subtotal']).number_format = '$ 00.00'
                        ws.cell(curr_row, 6, job['inv_subtotal']*0.2).number_format = '$ 00.00'
                        ws.cell(curr_row, 6).comment = Comment(job['summary'], "automation")
                    else:
                        ws.cell(curr_row, 5, job['open_est_subtotal']).number_format = '$ 00.00'
                    # 7
                    ws.cell(curr_row, 8, f"={get_column_letter(5)}{curr_row}-{get_column_letter(6)}{curr_row}-{get_column_letter(7)}{curr_row}").number_format = '$ 00.00'
                    ws.cell(curr_row, 9, job['payment_types'])
                    # 10 TODO: all doc checks complete
                    doc_check_complete_col = f'=IF(OR({", ".join([f"{get_column_letter(i)}{curr_row}=0" for i in range(11,20)])}), "N","Y")'
                    ws.cell(curr_row, 10, doc_check_complete_col).alignment = align_center
                    ws.cell(curr_row, 11, job['Before Photo'])
                    ws.cell(curr_row, 12, job['After Photo'])
                    ws.cell(curr_row, 13, job['Receipt Photo'])
                    ws.cell(curr_row, 14, job['Quote Description'])
                    ws.cell(curr_row, 15, job['Quote Signed'])
                    ws.cell(curr_row, 16, job['Quote Emailed'])
                    ws.cell(curr_row, 17, job['Invoice Description'])
                    ws.cell(curr_row, 18, job['Invoice Signed'])
                    ws.cell(curr_row, 19, job['Invoice Emailed'])
                    ws.cell(curr_row, 20, job['5 Star Review'])

                        # SUMIF(date_col, "date", profit/sales col)
                    
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
            

            # add 5 rows for any extras
            curr_row += 5

            # Fill in info 
            cat_row_info[cat] = (cat_row_start, curr_row-1)

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

            if cat == 'wk_complete_paid':
                formatted_cell(ws, curr_row, 11, f"=COUNT(K{cat_row_start}:K{curr_row-1})")
                formatted_cell(ws, curr_row, 12, f"=COUNT(L{cat_row_start}:L{curr_row-1})")
                formatted_cell(ws, curr_row, 13, f"=COUNT(M{cat_row_start}:M{curr_row-1})")
                formatted_cell(ws, curr_row, 14, f"=COUNT(N{cat_row_start}:N{curr_row-1})")
                formatted_cell(ws, curr_row, 15, f"=COUNT(O{cat_row_start}:O{curr_row-1})")
                formatted_cell(ws, curr_row, 16, f"=COUNT(P{cat_row_start}:P{curr_row-1})")
                formatted_cell(ws, curr_row, 17, f"=COUNT(Q{cat_row_start}:Q{curr_row-1})")
                formatted_cell(ws, curr_row, 18, f"=COUNT(R{cat_row_start}:R{curr_row-1})")
                formatted_cell(ws, curr_row, 19, f"=COUNT(S{cat_row_start}:S{curr_row-1})")
                formatted_cell(ws, curr_row, 20, f"=COUNT(T{cat_row_start}:T{curr_row-1})")

                formatted_cell(ws, curr_row+1, 11, f"=SUM(K{cat_row_start}:K{curr_row-1})")
                formatted_cell(ws, curr_row+1, 12, f"=SUM(L{cat_row_start}:L{curr_row-1})")
                formatted_cell(ws, curr_row+1, 13, f"=SUM(M{cat_row_start}:M{curr_row-1})")
                formatted_cell(ws, curr_row+1, 14, f"=SUM(N{cat_row_start}:N{curr_row-1})")
                formatted_cell(ws, curr_row+1, 15, f"=SUM(O{cat_row_start}:O{curr_row-1})")
                formatted_cell(ws, curr_row+1, 16, f"=SUM(P{cat_row_start}:P{curr_row-1})")
                formatted_cell(ws, curr_row+1, 17, f"=SUM(Q{cat_row_start}:Q{curr_row-1})")
                formatted_cell(ws, curr_row+1, 18, f"=SUM(R{cat_row_start}:R{curr_row-1})")
                formatted_cell(ws, curr_row+1, 19, f"=SUM(S{cat_row_start}:S{curr_row-1})")
                formatted_cell(ws, curr_row+1, 20, f"=SUM(T{cat_row_start}:T{curr_row-1})")
                curr_row += 1
            curr_row += 2
        
        # =SUMIF(B{cat_start}:B{cat_end}, date_str, E{cat_start}:E{cat_end}) -- sales
        # =SUMIF(B{cat_start}:B{cat_end}, date_str, H{cat_start}:H{cat_end}) -- profit
# monday_str

        days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday']
        profit_formulas = {day: '=' + ' + '.join([f'SUMIF(B{cat_row_info[cat][0]}:B{cat_row_info[cat][1]}, "{date_strs[day]}", H{cat_row_info[cat][0]}:H{cat_row_info[cat][1]})'for cat in cats_count_for_total]) for day in days}

        formatted_cell(ws, summary_top_row + 9,10, profit_formulas['monday'], font = font_bold, border = cell_border['bottomleft']) # These all rely on daily totals
        formatted_cell(ws, summary_top_row + 9,11, profit_formulas['tuesday'], font = font_bold, border = cell_border['bottom']) # These all rely on daily totals
        formatted_cell(ws, summary_top_row + 9,12, profit_formulas['wednesday'], font = font_bold, border = cell_border['bottom']) # These all rely on daily totals
        formatted_cell(ws, summary_top_row + 9,13, profit_formulas['thursday'], font = font_bold, border = cell_border['bottom']) # These all rely on daily totals
        formatted_cell(ws, summary_top_row + 9,14, profit_formulas['friday'], font = font_bold, border = cell_border['bottom']) # These all rely on daily totals
        formatted_cell(ws, summary_top_row + 9,16, profit_formulas['saturday'], font = font_green_bold, border = cell_border['bottomleft']) # These all rely on daily totals

        sales_formulas = {day: '=' + ' + '.join([f'SUMIF(B{cat_row_info[cat][0]}:B{cat_row_info[cat][1]}, "{date_strs[day]}", E{cat_row_info[cat][0]}:E{cat_row_info[cat][1]})'for cat in cats_count_for_total]) for day in days}

        formatted_cell(ws, summary_top_row + 17,10, sales_formulas['monday'], font = font_bold, border = cell_border['bottomleft']) # These all rely on daily totals
        formatted_cell(ws, summary_top_row + 17,11, sales_formulas['tuesday'], font = font_bold, border = cell_border['bottom'])#
        formatted_cell(ws, summary_top_row + 17,12, sales_formulas['wednesday'], font = font_bold, border = cell_border['bottom'])#
        formatted_cell(ws, summary_top_row + 17,13, sales_formulas['thursday'], font = font_bold, border = cell_border['bottom'])#
        formatted_cell(ws, summary_top_row + 17,14, sales_formulas['friday'], font = font_bold, border = cell_border['bottom'])#
        formatted_cell(ws, summary_top_row + 17,16, sales_formulas['saturday'], font = font_green_bold, border = cell_border['bottomleft']) # These all rely on daily totals

        count_success_formulas = {day: '=' + ' + '.join([f'COUNTIF(B{cat_row_info[cat][0]}:B{cat_row_info[cat][1]}, "{date_strs[day]}")'for cat in cats_count_for_total]) for day in days}

        formatted_cell(ws, summary_top_row + 10,10, count_success_formulas['monday']) # Actual count of monday jobs
        formatted_cell(ws, summary_top_row + 10,11, count_success_formulas['tuesday']) # Actual count of tuesday jobs
        formatted_cell(ws, summary_top_row + 10,12, count_success_formulas['wednesday']) # Actual count of wednesday jobs
        formatted_cell(ws, summary_top_row + 10,13, count_success_formulas['thursday']) # Actual count of thursday jobs
        formatted_cell(ws, summary_top_row + 10,14, count_success_formulas['friday']) # Actual count of friday jobs
        formatted_cell(ws, summary_top_row + 10,16, count_success_formulas['saturday']) # Actual count of saturday jobs

        count_unsuccess_formulas = {day: '=' + ' + '.join([f'COUNTIF(B{cat_row_info[cat][0]}:B{cat_row_info[cat][1]}, "{date_strs[day]}")'for cat in cats_count_for_unsuccessful]) for day in days}
        # Unsuccessful counts
        formatted_cell(ws, summary_top_row + 11,10, count_unsuccess_formulas['monday'])
        formatted_cell(ws, summary_top_row + 11,11, count_unsuccess_formulas['tuesday'])
        formatted_cell(ws, summary_top_row + 11,12, count_unsuccess_formulas['wednesday'])
        formatted_cell(ws, summary_top_row + 11,13, count_unsuccess_formulas['thursday'])
        formatted_cell(ws, summary_top_row + 11,14, count_unsuccess_formulas['friday'])
        formatted_cell(ws, summary_top_row + 11,16, count_unsuccess_formulas['saturday'])

        end_of_comp_paid = cat_row_info['wk_complete_paid'][1]
        formatted_cell(ws, summary_top_row + 3,10, f'=K{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,11, f'=L{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,12, f'=M{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,13, f'=N{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,14, f'=O{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,15, f'=P{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,16, f'=Q{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,17, f'=R{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        formatted_cell(ws, summary_top_row + 3,18, f'=S{end_of_comp_paid+2}', font = font_bold, border = cell_border['left'])
        
        
        review_count_formula = '=' + ' + '.join([f'SUM(T{cat_row_info[cat][0]}:T{cat_row_info[cat][1]})'for cat in cats_count_for_total])
        formatted_cell(ws, summary_top_row + 3,19, review_count_formula, font = font_bold, border = cell_border['leftright']) # total 5 star reviews

        formatted_cell(ws, summary_top_row + 4,10, f'=K{end_of_comp_paid+2}/K{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,11, f'=L{end_of_comp_paid+2}/L{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,12, f'=M{end_of_comp_paid+2}/M{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,13, f'=N{end_of_comp_paid+2}/N{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,14, f'=O{end_of_comp_paid+2}/O{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,15, f'=P{end_of_comp_paid+2}/P{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,16, f'=Q{end_of_comp_paid+2}/Q{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,17, f'=R{end_of_comp_paid+2}/R{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])
        formatted_cell(ws, summary_top_row + 4,18, f'=S{end_of_comp_paid+2}/S{end_of_comp_paid+1}', font = font_bold, border = cell_border['bottomleft'])

        

        wk_profit_awaiting_formula = '=' + ' + '.join([f'SUM(H{cat_row_info[cat][0]}:H{cat_row_info[cat][1]})'for cat in cats_count_awaiting_pay_wk])
        formatted_cell(ws, summary_top_row + 9,17, wk_profit_awaiting_formula, font = font_red_bold, border = cell_border['bottomleft'])
        wkend_profit_awaiting_formula = '=' + ' + '.join([f'SUM(H{cat_row_info[cat][0]}:H{cat_row_info[cat][1]})'for cat in cats_count_awaiting_pay_wkend])
        formatted_cell(ws, summary_top_row + 9,18, wkend_profit_awaiting_formula, font = font_red_bold, border = cell_border['bottomleftright'])

        #mngment 
        wk_sales_awaiting_formula = '=' + ' + '.join([f'SUM(E{cat_row_info[cat][0]}:E{cat_row_info[cat][1]})'for cat in cats_count_awaiting_pay_wk])
        formatted_cell(ws, summary_top_row + 17,17, wk_sales_awaiting_formula, font = font_red_bold, border = cell_border['bottomleft'])
        wkend_sales_awaiting_formula = '=' + ' + '.join([f'SUM(E{cat_row_info[cat][0]}:E{cat_row_info[cat][1]})'for cat in cats_count_awaiting_pay_wkend])
        formatted_cell(ws, summary_top_row + 17,18, wkend_sales_awaiting_formula, font = font_red_bold, border = cell_border['bottomleftright'])
        
        # potential

        wk_profit_potential_formula = '=' + ' + '.join([f'SUM(H{cat_row_info[cat][0]}:H{cat_row_info[cat][1]})'for cat in cats_count_for_potential_wk])
        formatted_cell(ws, summary_top_row + 7,6, wk_profit_potential_formula, font = font_red, border = cell_border['topright'])
        # wkend_profit_potential_formula = '=' + ' + '.join([f'SUM(H{cat_row_info[cat][0]}:H{cat_row_info[cat][1]})'for cat in cats_count_for_potential_wkend])
        formatted_cell(ws, summary_top_row + 10,6, border = cell_border['right'])
        



        for row in ws[f'V{27}:X{curr_row-3}']:
            for cell in row:
                cell.border = cell_border_full


    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()