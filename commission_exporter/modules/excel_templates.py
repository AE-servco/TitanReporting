import datetime as dt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule

import modules.helpers as helpers


# function for each "box" in the summary
# function for the daily summary bits, one for monthly, one for weekly.
# function for actual jobs
#   - function for each job line maybe?
# calculate threshold function
#   - based on public holidays and days per week/month otherwise

class CommissionSpreadSheetExporter:
    def __init__(self, jobs_by_tech: dict[str, list[dict]], end_date: dt.date, timeframe: str):
        self.jobs_by_tech = jobs_by_tech
        self.end_date = end_date
        self.timeframe = timeframe
        # self.first_sheet_created = False
        self.curr_worksheet = None
        self.curr_row = 1

        # =========================== STYLING ===========================
        self.sales_color = '3da813'
        self.installer_color = 'e3b019'
        self.unknown_color = '8093f2'

        thin_border = Side(style='thin', color='000000') # Black thin border
        med_border = Side(style='medium', color='000000') # Black medium border
        thick_border = Side(style='thick', color='000000') # Black thick border
        double_border = Side(style='double', color='000000')

        self.yellow_fill = PatternFill(patternType='solid', fgColor='FFFF00')
        self.blue_fill = PatternFill(patternType='solid', fgColor='3399F2')

        self.font_red = Font(color='FF0000')
        self.font_red_bold = Font(color='FF0000', bold=True)
        self.font_green_bold = Font(color='1F9C40', bold=True)
        self.font_green = Font(color='1F9C40')
        self.font_purple = Font(color='5F4F7D')
        self.font_bold = Font(bold=True)

        self.cell_border_full = Border(top=thin_border, bottom=thin_border, right=thin_border, left=thin_border)
        self.border_double_bottom = Border(bottom=double_border)

        self.accounting_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
        self.percentage_format = '0.00%'

        self.cell_border = {
            'topleft': Border(left=med_border, top=med_border),
            'topright': Border(right=med_border, top=med_border),
            'topleftright': Border(left=med_border, top=med_border, right=med_border),
            'bottomleft': Border(left=med_border, bottom=med_border),
            'bottomright': Border(right=med_border, bottom=med_border),
            'bottomtop': Border(top=med_border, bottom=med_border),
            'bottomtop_double': Border(top=med_border, bottom=double_border),
            'bottomtopright': Border(top=med_border, bottom=med_border, right=med_border),
            'bottomleftright': Border(left=med_border, bottom=med_border, right=med_border),
            'bottom': Border(bottom=med_border),
            'top': Border(top=med_border),
            'right': Border(right=med_border),
            'left': Border(left=med_border),
            'leftright': Border(left=med_border, right=med_border),
        }

        self.align_center = Alignment(horizontal='center')

        # =========================== CONSTANTS ===========================
        self.CATEGORY_ORDER = {
            'wk_complete_paid': 'COMPLETED & PAID JOBS', 
            'wkend_complete_paid': 'WEEKEND COMPLETED & PAID JOBS', 
            'ah_complete_paid': 'AFTERHOURS COMPLETED & PAID JOBS', 
            'prev': 'PREVIOUS JOBS COMPLETED & PAID (COMMISSION) - Modoras Team Please do ADD to THIS SECTION or AMEND/TOUCH', 
            'wk_complete_unpaid': 'CURRENT JOBS COMPLETED (AWAITING PAYMENT)', 
            'wkend_complete_unpaid': 'WEEKEND CURRENT JOBS COMPLETED (AWAITING PAYMENT)', 
            'ah_complete_unpaid': 'AFTERHOURS CURRENT JOBS COMPLETED (AWAITING PAYMENT)', 
            'wk_wo': 'CURRENT WORK ORDERS (AWAITING PAYMENT)', 
            'wkend_wo': 'WEEKEND WORK ORDERS (AWAITING PAYMENT)', 
            'ah_wo': 'AFTERHOURS WORK ORDERS (AWAITING PAYMENT)', 
            'wk_unsuccessful': 'UNSUCCESSFUL JOBS',
            'wkend_unsuccessful': 'WEEKEND UNSUCCESSFUL JOBS',
            'ah_unsuccessful': 'AFTERHOURS UNSUCCESSFUL JOBS',
            'wk_uncategorised': 'WEEK UNCATEGORISED',
            'wkend_uncategorised': 'WEEKEND UNCATEGORISED',
            'ah_uncategorised': 'AFTERHOURS UNCATEGORISED',
        }

    def formatted_cell(worksheet, row: int, col: int, val = None, font: Font | None=None, border: Border | None=None, number_format: str | None=None, fill: PatternFill | None=None):
        if val or val == 0:
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

    def title_row(self):

        pass

    def job_count_box(self):
        pass

    def profit_target_box(self):
        pass

    def payout_box(self):
        pass

    def doc_check_count_box(self):
        pass

    def day_summaries_monthly(self):
        pass

    def day_summaries_weekly(self):
        pass

    def job_section_title(self):
        pass
    
    def put_job_row(self, job: dict):
        ws = self.curr_worksheet
        if job['first_appt_start_str'] != curr_date:
            job_border = self.cell_border['top']
        else:
            job_border = None
        curr_date = job['first_appt_start_str']

        if job['complaint_tag_present']:
            self.formatted_cell(ws, curr_row, col_offset, 'COMPLAINT', font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 1, job_count, font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 2, job['first_appt_start_str'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 3, int(job['num']), font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 4, job['suburb'], font=cat_font, border = job_border)
        if not job['unsuccessful']:
            self.formatted_cell(ws, curr_row, col_offset + 5, job['inv_subtotal'], font=cat_font, border = job_border, number_format=accounting_format)
            self.formatted_cell(ws, curr_row, col_offset + 6, job['inv_subtotal']*0.2, font=cat_font, border = job_border, number_format=accounting_format)
            self.formatted_cell(ws, curr_row, col_offset + 6, font=cat_font, border = job_border).comment = Comment(job['summary'], "automation")
        else:
            self.formatted_cell(ws, curr_row, col_offset + 5, job['open_est_subtotal'], font=cat_font, border = job_border, number_format=accounting_format)
            self.formatted_cell(ws, curr_row, col_offset + 6, "", font=cat_font, border = job_border)
        # 7
        self.formatted_cell(ws, curr_row, col_offset + 7, "", font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 8, f"={get_column_letter(col_offset + 5)}{curr_row}-{get_column_letter(col_offset + 6)}{curr_row}-{get_column_letter(col_offset + 7)}{curr_row}", font=cat_font, border = job_border, number_format=accounting_format)
        self.formatted_cell(ws, curr_row, col_offset + 9, job['payment_types'], font=cat_font, border = job_border)
        # 10 TODO: all doc checks complete
        doc_check_complete_col = f'=IF(OR({", ".join([f"{get_column_letter(col_offset + i)}{curr_row}=0" for i in range(11,20)])}), "N","Y")'
        self.formatted_cell(ws, curr_row, col_offset + 10, doc_check_complete_col, font=cat_font, border = job_border).alignment = align_center
        self.formatted_cell(ws, curr_row, col_offset + 11, job['Before Photo'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 12, job['After Photo'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 13, job['Receipt Photo'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 14, job['Quote Description'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 15, job['Quote Signed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 16, job['Quote Emailed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 17, job['Invoice Description'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 18, job['Invoice Signed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 19, job['Invoice Emailed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, curr_row, col_offset + 20, job['5 Star Review'], font=cat_font, border = job_border)

        self.formatted_cell(ws, curr_row, col_offset + 25, f"=ROUND(F{curr_row}*1.1,2)", font=font_purple)
        self.formatted_cell(ws, curr_row, col_offset + 26, f"=ROUND(Z{curr_row} - W{curr_row} - X{curr_row} - Y{curr_row},2)", font=font_purple)
        
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
                self.formatted_cell(ws, curr_row, col_offset + 22, f"={'+'.join(p_types['EF'] + p_types['Cr'])}", font=cat_font, number_format=accounting_format)
            if p_types['Ca']:
                self.formatted_cell(ws, curr_row, col_offset + 23, f"={'+'.join(p_types['Ca'])}", font=cat_font, number_format=accounting_format)
        curr_row += 1
        job_count += 1
        pass
    
    def put_job_category(self, cat: str, cat_text: str):
        pass

    def extra_formatting(self):
        pass

    def build_sheet(self, wb: Workbook, tech: str | int):
        job_cats = self.jobs_by_tech[tech]
        if not self.curr_worksheet:
            self.curr_worksheet = wb.active
            self.curr_worksheet.title = tech[:-1]
            self.first_sheet_created = True
        else:
            self.curr_worksheet = wb.create_sheet(title=tech[:-1])

        if tech[-1] == 'S':
            self.curr_worksheet.sheet_properties.tabColor = self.sales_color
        elif tech[-1] == 'I':
            self.curr_worksheet.sheet_properties.tabColor = self.installer_color
        else:
            self.curr_worksheet.sheet_properties.tabColor = self.unknown_color

        summary_top_row = 1

        # ========================= TOP SECTION - SUMMARIES =========================
        # Will need to call some of these after doing jobs because they will need to know cat start and ends etc.
        self.title_row()
        self.job_count_box()
        self.profit_target_box()
        self.payout_box()
        self.doc_check_count_box()
        if self.timeframe == 'weekly':
            self.day_summaries_weekly()
        elif self.timeframe == 'monthly':
            self.day_summaries_monthly()
        else:
            raise ValueError("Timeframe must be 'weekly' or 'monthly'")
        
        # ========================== BOTTOM SECTION - JOBS ==========================
        self.job_section_title()

        for cat, cat_text in self.CATEGORY_ORDER.items():
            self.put_job_category(cat, cat_text)
        

    def build_workbook(self):
        # Final function that actually builds everything
        wb = Workbook()
        col_offset = 1

        pass