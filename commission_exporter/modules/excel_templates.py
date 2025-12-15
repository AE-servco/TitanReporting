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
    def __init__(self, jobs_by_tech: dict[str, list[dict]], end_date: dt.date, timeframe: str, col_offset: int):
        self.jobs_by_tech = jobs_by_tech
        self.end_date = end_date
        self.timeframe = timeframe
        self.curr_worksheet = None
        self.curr_row = 1
        self.bottom_row = 1
        self.job_count = 1
        self.col_offset = col_offset
        self.cat_row_info = {}
        self.threshold_day_num = 0
        self.curr_date = ""

        self.cats_count_for_total = [
            'wk_complete_paid',
            'wkend_complete_paid',
            'wk_complete_unpaid',
            'wkend_complete_unpaid',
            'wk_wo',
            'wkend_wo',
            'ah_complete_paid',
            'ah_complete_unpaid',
            'ah_wo',
        ]

        self.cats_count_awaiting_pay_wk = [
            'wk_complete_unpaid',
            'wk_wo',
        ]

        self.cats_count_awaiting_pay_wkend = [
            'wkend_complete_unpaid',
            'wkend_wo',
            'ah_complete_unpaid',
            'ah_wo',
        ]

        self.cats_count_for_potential_wk = [
            'wk_complete_paid',
            'wk_complete_unpaid',
            'wk_wo',
            'wk_unsuccessful',
        ]

        self.cats_count_for_potential_wkend = [
            'wkend_complete_paid',
            'wkend_complete_unpaid',
            'wkend_wo',
            'wkend_unsuccessful',
        ]

        self.cats_count_for_unsuccessful = [
            'wk_unsuccessful',
            'wkend_unsuccessful',
            'ah_unsuccessful',
        ]

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
            'topthin': Border(top=thin_border),
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

    def formatted_cell(self, worksheet: Worksheet, row: int, col: int, val = None, font: Font | None=None, border: Border | None=None, number_format: str | None=None, fill: PatternFill | None=None):
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

    def title_row(self, start_row: int = 1):
        ws = self.curr_worksheet
        col_offset = self.col_offset
        self.curr_row = start_row

        self.formatted_cell(ws, self.curr_row, col_offset + 1, 'WEEKLY COMMISSION', font = self.font_bold)
        self.formatted_cell(ws, self.curr_row, col_offset + 5, 'WEEK ENDING', font = self.font_red_bold)
        self.formatted_cell(ws, self.curr_row, col_offset + 6, self.end_date.strftime("%d/%m/%Y"), font = self.font_red_bold)
        return

    def job_count_box(self, start_row: int):
        # THIS IS MONTHLY ===============================
        ws = self.curr_worksheet
        col_offset = self.col_offset
        self.formatted_cell(ws, start_row, col_offset + 1, 'BOOKED JOBS',font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 2, '=C4+C5', font = self.font_bold, border = self.cell_border['topright'])
        self.formatted_cell(ws, start_row + 1, col_offset + 1, 'SUCCESSFUL', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 1, 'UNSUCCESSFUL', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 3, col_offset + 1, 'SUCCESSFUL (%)', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 3, col_offset + 2, '=C4/(C4+C5)', font = self.font_bold, border = self.cell_border['right'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 5, col_offset + 1, 'AVERAGE SALE', font = self.font_bold, border = self.cell_border['bottomleft'])
        # self.formatted_cell(ws, start_row + 11, col_offset + 17, '=SUM(K11:O11) + SUM(K17:O17) + SUM(K23:O23) + SUM(K29:O29) + SUM(K35:O35) + SUM(K41:O41)', font = font_bold, border = cell_border['bottomleft'], number_format=accounting_format)
        if self.timeframe == 'monthly':
            self.formatted_cell(ws, start_row + 1, col_offset + 2, '=SUM(K12:O12) + SUM(K18:O18) + SUM(K24:O24) + SUM(K30:O30) + SUM(K36:O36) + SUM(K42:O42)', font = self.font_bold, border = self.cell_border['right'])
            self.formatted_cell(ws, start_row + 2, col_offset + 2, '=SUM(K13:O13) + SUM(K19:O19) + SUM(K25:O25) + SUM(K31:O31) + SUM(K37:O37) + SUM(K43:O43)', font = self.font_bold, border = self.cell_border['right'])
            self.formatted_cell(ws, start_row + 5, col_offset + 2, '=R12/C4', font = self.font_bold, border = self.cell_border['bottomright'], number_format=self.accounting_format)
        elif self.timeframe == 'weekly':
            # TODO
            pass

        self.formatted_cell(ws, start_row + 4, col_offset + 1, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 4, col_offset + 2, border = self.cell_border['right'])

        self.curr_row = start_row + 5
        if self.bottom_row < self.curr_row:
            self.bottom_row = self.curr_row
        return

    def profit_target_box(self, start_row: int):
        ws = self.curr_worksheet
        col_offset = self.col_offset

        self.formatted_cell(ws, start_row, col_offset + 4, 'PROFIT TARGET', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 6, self.threshold_day_num * 5000, font = self.font_bold, border = self.cell_border['topleft'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row, col_offset + 5, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 6, border = self.cell_border['topright'])
        self.formatted_cell(ws, start_row + 1, col_offset + 4, 'Tier 1', border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 1, col_offset + 5, f'<${self.threshold_day_num * 5000}')
        self.formatted_cell(ws, start_row + 1, col_offset + 6, '=0.05', border = self.cell_border['right'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 2, col_offset + 4, 'Tier 2', border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 2, col_offset + 5, f'>=${self.threshold_day_num * 5000}', border = self.cell_border['bottom'])
        self.formatted_cell(ws, start_row + 2, col_offset + 6, '=0.1', border = self.cell_border['bottomright'], number_format=self.percentage_format)

        self.curr_row = start_row + 2
        if self.bottom_row < self.curr_row:
            self.bottom_row = self.curr_row
        return

    def payout_box(self, start_row: int):
        # TODO: make the letters dynamic in here
        # Some monthly things here (R12)
        ws = self.curr_worksheet
        col_offset = self.col_offset

        self.formatted_cell(ws, start_row, col_offset + 5, 'ACTUAL', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 6, 'POTENTIAL', font = self.font_bold, border = self.cell_border['topright'])
        self.formatted_cell(ws, start_row, col_offset + 7, 'Exc. SUPER', font = self.font_green_bold)
        self.formatted_cell(ws, start_row + 1, col_offset + 4, 'NET PROFIT', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row + 2, col_offset + 4, 'UNLOCKED', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 5, f'=IF(R12>={self.threshold_day_num * 5000},G5,G4)', border = self.cell_border['left'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 4, 'COMMISSION - PAY OUT', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 3, col_offset + 5, '=F8*F9', border = self.cell_border['left'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 7, '=F10/1.12', font = self.font_green_bold, number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 4, col_offset + 4, 'EMERGENCY', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 4, col_offset + 5, '=S12', border = self.cell_border['left'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 5, col_offset + 4, 'EMERGENCY - PAY OUT', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 5, col_offset + 5, '=F11*0.25', border = self.cell_border['left'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 5, col_offset + 7, '=F12/1.12', font = self.font_green_bold, number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 6, col_offset + 4, 'PREV. WEEK', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 6, col_offset + 5, 0, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 7, col_offset + 4, 'PREV. WEEK - PAY OUT', font = self.font_bold, border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 7, col_offset + 5, '=F13*0.05', border = self.cell_border['bottomleft'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 7, col_offset + 7, '=F14/1.12', font = self.font_green_bold, number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 8, col_offset + 4, '5 Star Review', font = self.font_green_bold, border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 8, col_offset + 5, '=F16*50', border = self.cell_border['left'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 9, col_offset + 4, '5 Star Notes', font = self.font_green_bold, border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 9, col_offset + 5, '=T4', border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 6, f'=IF(G8>={self.threshold_day_num * 5000},G5,G4)', font=self.font_red, border = self.cell_border['right'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 6, '=G8*G9', font = self.font_red, border = self.cell_border['right'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 5, col_offset + 6, border = self.cell_border['right'])
        self.formatted_cell(ws, start_row + 6, col_offset + 6, 0, font = self.font_red, border = self.cell_border['right'])
        self.formatted_cell(ws, start_row + 7, col_offset + 6, '==G13*0.05', font = self.font_red, border = self.cell_border['bottomright'], number_format=self.accounting_format)
        
        # Subtracting red from payout
        subtract_red_formula_wk = f'SUMIF(K{self.cat_row_info["wk_complete_paid"][0]}:K{self.cat_row_info["wk_complete_paid"][1]}, "N", I{self.cat_row_info["wk_complete_paid"][0]}:I{self.cat_row_info["wk_complete_paid"][1]})'
        self.formatted_cell(ws, start_row + 1, col_offset + 5, f'=R12-R10-{subtract_red_formula_wk}', border = self.cell_border['topleft'], number_format=self.accounting_format)
        
        wk_profit_potential_formula = '=' + ' + '.join([f'SUM(I{self.cat_row_info[cat][0]}:I{self.cat_row_info[cat][1]})'for cat in self.cats_count_for_potential_wk])
        self.formatted_cell(ws, start_row + 1, col_offset + 6, wk_profit_potential_formula, font = self.font_red, border = self.cell_border['topright'], number_format=self.accounting_format)
        # wkend_profit_potential_formula = '=' + ' + '.join([f'SUM(I{cat_row_info[cat][0]}:I{cat_row_info[cat][1]})'for cat in cats_count_for_potential_wkend])
        self.formatted_cell(ws, start_row + 4, col_offset + 6, border = self.cell_border['right'])
        
        self.curr_row = start_row + 9
        if self.bottom_row < self.curr_row:
            self.bottom_row = self.curr_row
        return

    def doc_check_count_box(self, start_row: int):
        ws = self.curr_worksheet
        col_offset = self.col_offset

        self.formatted_cell(ws, start_row, col_offset + 10, 'PHOTOS', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 11, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 12, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 13, 'QUOTE', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 14, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 15, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 16, 'INVOICE', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 17, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 18, border = self.cell_border['topright'])
        self.formatted_cell(ws, start_row, col_offset + 19, border = self.cell_border['topright'])
        self.formatted_cell(ws, start_row + 1, col_offset + 10, 'BEFORE', font = self.font_bold, border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 1, col_offset + 11, 'AFTER', font = self.font_bold, border = self.cell_border['bottom'])
        self.formatted_cell(ws, start_row + 1, col_offset + 12, 'RECEIPT', font = self.font_bold, border = self.cell_border['bottom'])
        self.formatted_cell(ws, start_row + 1, col_offset + 13, 'DESCRIPTION', font = self.font_bold, border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 1, col_offset + 14, 'SIGNED', font = self.font_bold, border = self.cell_border['bottom'])
        self.formatted_cell(ws, start_row + 1, col_offset + 15, 'EMAILED', font = self.font_bold, border = self.cell_border['bottom'])
        self.formatted_cell(ws, start_row + 1, col_offset + 16, 'DESCRIPTION', font = self.font_bold, border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 1, col_offset + 17, 'SIGNED', font = self.font_bold, border = self.cell_border['bottom'])
        self.formatted_cell(ws, start_row + 1, col_offset + 18, 'EMAILED', font = self.font_bold, border = self.cell_border['bottom'])
        self.formatted_cell(ws, start_row + 1, col_offset + 19, '5 Star Review', font = self.font_bold, border = self.cell_border['bottomleftright'])
        self.formatted_cell(ws, start_row + 2, col_offset + 9, 'TAKEN', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row + 3, col_offset + 9, '%', font = self.font_bold, border = self.cell_border['bottomleft'])
        self.formatted_cell(ws, start_row + 3, col_offset + 19, border = self.cell_border['bottomleftright'])
        
        # TODO: can probably make this dynamic and a for loop
        end_of_comp_paid = self.cat_row_info['wk_complete_paid'][1]
        self.formatted_cell(ws, start_row + 2, col_offset + 10, f'=L{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 11, f'=M{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 12, f'=N{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 13, f'=O{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 14, f'=P{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 15, f'=Q{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 16, f'=R{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 17, f'=S{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        self.formatted_cell(ws, start_row + 2, col_offset + 18, f'=T{end_of_comp_paid+2}', font = self.font_bold, border = self.cell_border['left'])
        
        review_count_formula = '=' + ' + '.join([f'SUM(U{self.cat_row_info[cat][0]}:U{self.cat_row_info[cat][1]})'for cat in self.cats_count_for_total])
        self.formatted_cell(ws, start_row + 2, col_offset + 19, review_count_formula, font = self.font_bold, border = self.cell_border['leftright']) # total 5 star reviews

        self.formatted_cell(ws, start_row + 3, col_offset + 10, f'=L{end_of_comp_paid+2}/L{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 11, f'=M{end_of_comp_paid+2}/M{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 12, f'=N{end_of_comp_paid+2}/N{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 13, f'=O{end_of_comp_paid+2}/O{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 14, f'=P{end_of_comp_paid+2}/P{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 15, f'=Q{end_of_comp_paid+2}/Q{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 16, f'=R{end_of_comp_paid+2}/R{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 17, f'=S{end_of_comp_paid+2}/S{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        self.formatted_cell(ws, start_row + 3, col_offset + 18, f'=T{end_of_comp_paid+2}/T{end_of_comp_paid+1}', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.percentage_format)
        
        
        self.curr_row = start_row + 3
        if self.bottom_row < self.curr_row:
            self.bottom_row = self.curr_row
        return

    def day_summaries_monthly(self, start_row: int):
        ws = self.curr_worksheet
        col_offset = self.col_offset

        dates_in_month = helpers.get_dates_in_month_datetime(self.end_date.year, self.end_date.month)

        self.formatted_cell(ws, start_row, col_offset + 10, 'DAILY NET PROFIT', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 11, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 12, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 13, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 14, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 15, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 16, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 17, 'WEEKDAY', font = self.font_red_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 18, 'WEEKEND', font = self.font_red_bold, border = self.cell_border['topleftright'])
        self.formatted_cell(ws, start_row + 1, col_offset + 10, 'MON', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 11, 'TUE', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 12, 'WED', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 13, 'THU', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 14, 'FRI', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 15, 'SAT', font = self.font_green_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 16, 'SUN', font = self.font_green_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 17, 'Awaiting Payment', font = self.font_red_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row + 1, col_offset + 18, 'Awaiting Payment', font = self.font_red_bold, border = self.cell_border['topleftright'])
        self.formatted_cell(ws, start_row + 3, col_offset + 17, 'Total Payable', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row + 3, col_offset + 18, 'Total Payable', font = self.font_bold, border = self.cell_border['topleftright'])
        self.formatted_cell(ws, start_row + 4, col_offset + 17, '=SUM(K11:O11) + SUM(K17:O17) + SUM(K23:O23) + SUM(K29:O29) + SUM(K35:O35) + SUM(K41:O41)', font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 4, col_offset + 18, '=SUM(P11:Q11) + SUM(P17:Q17) + SUM(P23:Q23) + SUM(P29:Q29) + SUM(P35:Q35) + SUM(P41:Q41)', font = self.font_bold, border = self.cell_border['bottomright'], number_format=self.accounting_format)
        self.formatted_cell(ws, start_row + 4, col_offset + 9, 'Successful')
        self.formatted_cell(ws, start_row + 5, col_offset + 9, 'Unsuccessful')
        self.formatted_cell(ws, start_row + 6, col_offset + 9, 'Success rate')
        self.formatted_cell(ws, start_row + 7, col_offset + 9, 'Avg sale')

        profit_formulas = {day: '=' + ' + '.join([f'SUMIF(C{self.cat_row_info[cat][0]}:C{self.cat_row_info[cat][1]}, "{day.strftime("%d/%m/%Y")}", I{self.cat_row_info[cat][0]}:I{self.cat_row_info[cat][1]})'for cat in self.cats_count_for_total]) for day in dates_in_month}
        count_success_formulas = {day: '=' + ' + '.join([f'COUNTIF(C{self.cat_row_info[cat][0]}:C{self.cat_row_info[cat][1]}, "{day.strftime("%d/%m/%Y")}")'for cat in self.cats_count_for_total]) for day in dates_in_month}
        count_unsuccess_formulas = {day: '=' + ' + '.join([f'COUNTIF(C{self.cat_row_info[cat][0]}:C{self.cat_row_info[cat][1]}, "{day.strftime("%d/%m/%Y")}")'for cat in self.cats_count_for_unsuccessful]) for day in dates_in_month}

        day_start_row = start_row + 2
        reset_col = 10
        start_col = reset_col

        SUMMARY_COL_LENGTH = 0
        for day in dates_in_month:
            day_row = day_start_row
            day_of_week = day.weekday()
            if day_of_week == 0:
                # If monday, write the LHS words 
                self.formatted_cell(ws, day_row + 2, col_offset + start_col - 1, 'Successful')
                self.formatted_cell(ws, day_row + 2 + 1, col_offset + start_col - 1, 'Unsuccessful')
                self.formatted_cell(ws, day_row + 2 + 2, col_offset + start_col - 1, 'Success rate')
                self.formatted_cell(ws, day_row + 2 + 3, col_offset + start_col - 1, 'Avg sale')
            
            summary_font = self.font_bold
            if day_of_week in [5,6]:
                summary_font = self.font_green_bold

            curr_col = col_offset + start_col + day_of_week
            self.formatted_cell(ws, day_row, curr_col, day.strftime("%d/%m"), font = summary_font, border = self.cell_border['top']) 
            day_row += 1
            self.formatted_cell(ws, day_row, curr_col, profit_formulas[day], font = summary_font, number_format=self.accounting_format) 
            day_row += 1
            self.formatted_cell(ws, day_row, curr_col, count_success_formulas[day]) 
            day_row += 1
            self.formatted_cell(ws, day_row, curr_col, count_unsuccess_formulas[day])
            day_row += 1
            curr_col_letter = get_column_letter(curr_col)
            self.formatted_cell(ws, day_row, curr_col, f'={curr_col_letter}{day_row-2}/({curr_col_letter}{day_row-2}+{curr_col_letter}{day_row-1})', number_format=self.percentage_format)
            day_row += 1
            self.formatted_cell(ws, day_row, curr_col, f'={curr_col_letter}{day_row-4}/{curr_col_letter}{day_row-3}', number_format=self.accounting_format)
            day_row += 1

            SUMMARY_COL_LENGTH = day_row - day_start_row
            if day_of_week == 6:
                # If sunday, push row to next block
                day_start_row += SUMMARY_COL_LENGTH

        wk_profit_awaiting_formula = '=' + ' + '.join([f'SUM(I{self.cat_row_info[cat][0]}:I{self.cat_row_info[cat][1]})'for cat in self.cats_count_awaiting_pay_wk])
        self.formatted_cell(ws, start_row + 2, col_offset + 17, wk_profit_awaiting_formula, font = self.font_red_bold, border = self.cell_border['bottomleft'], number_format=self.accounting_format)

        wkend_profit_awaiting_formula = '=' + ' + '.join([f'SUM(I{self.cat_row_info[cat][0]}:I{self.cat_row_info[cat][1]})'for cat in self.cats_count_awaiting_pay_wkend])
        self.formatted_cell(ws, start_row + 2, col_offset + 18, wkend_profit_awaiting_formula, font = self.font_red_bold, border = self.cell_border['bottomleftright'], number_format=self.accounting_format)

        return

    def day_summaries_weekly(self, start_row: int):
#         =SUMIF(B{cat_start}:B{cat_end}, date_str, E{cat_start}:E{cat_end}) -- sales
#         =SUMIF(B{cat_start}:B{cat_end}, date_str, H{cat_start}:H{cat_end}) -- profit
# monday_str
        ws = self.curr_worksheet
        col_offset = self.col_offset

        dates = {
            'monday': self.end_date - dt.timedelta(days=6),
            'tuesday': self.end_date - dt.timedelta(days=5),
            'wednesday': self.end_date - dt.timedelta(days=4),
            'thursday': self.end_date - dt.timedelta(days=3),
            'friday': self.end_date - dt.timedelta(days=2),
            'saturday': self.end_date - dt.timedelta(days=1),
            'sunday': self.end_date
        }
    
        date_strs = {day: date.strftime("%d/%m/%Y") for day, date in dates.items()}

        self.formatted_cell(ws, start_row, col_offset + 10, 'DAILY NET PROFIT', font = self.font_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 11, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 12, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 13, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 14, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 15, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 16, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 17, border = self.cell_border['top'])
        self.formatted_cell(ws, start_row, col_offset + 18, 'WEEKDAY', font = self.font_red_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row, col_offset + 19, 'WEEKEND', font = self.font_red_bold, border = self.cell_border['topleftright'])
        self.formatted_cell(ws, start_row + 1, col_offset + 10, 'MON', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 11, 'TUE', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 12, 'WED', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 13, 'THU', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 14, 'FRI', font = self.font_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 15, 'TOTAL', font = self.font_bold, border = self.cell_border['bottomtop_double'])

        self.formatted_cell(ws, start_row + 1, col_offset + 16, 'SAT', font = self.font_green_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 17, 'SUN', font = self.font_green_bold, border = self.cell_border['bottomtop_double'])
        self.formatted_cell(ws, start_row + 1, col_offset + 18, 'Awaiting Payment', font = self.font_red_bold, border = self.cell_border['topleft'])
        self.formatted_cell(ws, start_row + 1, col_offset + 19, 'Awaiting Payment', font = self.font_red_bold, border = self.cell_border['topleftright'])

        self.formatted_cell(ws, start_row + 3, col_offset + 9, 'Successful')
        self.formatted_cell(ws, start_row + 4, col_offset + 9, 'Unsuccessful')
        self.formatted_cell(ws, start_row + 5, col_offset + 9, 'Success rate')
        self.formatted_cell(ws, start_row + 6, col_offset + 9, 'Avg sale')

        days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
        profit_formulas = {day: '=' + ' + '.join([f'SUMIF(C{self.cat_row_info[cat][0]}:C{self.cat_row_info[cat][1]}, "{date_strs[day]}", I{self.cat_row_info[cat][0]}:I{self.cat_row_info[cat][1]})'for cat in self.cats_count_for_total]) for day in days}

        self.formatted_cell(ws, start_row + 2, col_offset + 10, profit_formulas['monday'], font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.accounting_format) # These all rely on daily totals
        self.formatted_cell(ws, start_row + 2, col_offset + 11, profit_formulas['tuesday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format) # These all rely on daily totals
        self.formatted_cell(ws, start_row + 2, col_offset + 12, profit_formulas['wednesday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format) # These all rely on daily totals
        self.formatted_cell(ws, start_row + 2, col_offset + 13, profit_formulas['thursday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format) # These all rely on daily totals
        self.formatted_cell(ws, start_row + 2, col_offset + 14, profit_formulas['friday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format) # These all rely on daily totals

        total_wk_profit_formula = '=' + '+'.join([f'SUM({get_column_letter(col_offset + 10 + i)}{start_row + 2})' for i in range(5)])
        self.formatted_cell(ws, start_row + 2, col_offset + 15, total_wk_profit_formula, font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format) # These all rely on daily totals

        self.formatted_cell(ws, start_row + 2, col_offset + 16, profit_formulas['saturday'], font = self.font_green_bold, border = self.cell_border['bottomleft'], number_format=self.accounting_format) 
        self.formatted_cell(ws, start_row + 2, col_offset + 17, profit_formulas['sunday'], font = self.font_green_bold, border = self.cell_border['bottomright'], number_format=self.accounting_format) 

        # sales_formulas = {day: '=' + ' + '.join([f'SUMIF(C{self.cat_row_info[cat][0]}:C{self.cat_row_info[cat][1]}, "{date_strs[day]}", F{self.cat_row_info[cat][0]}:F{self.cat_row_info[cat][1]})'for cat in self.cats_count_for_total]) for day in days}

        # self.formatted_cell(ws, start_row + 17, col_offset + 10, sales_formulas['monday'], font = self.font_bold, border = self.cell_border['bottomleft'], number_format=self.accounting_format) # These all rely on daily totals
        # self.formatted_cell(ws, start_row + 17, col_offset + 11, sales_formulas['tuesday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format)#
        # self.formatted_cell(ws, start_row + 17, col_offset + 12, sales_formulas['wednesday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format)#
        # self.formatted_cell(ws, start_row + 17, col_offset + 13, sales_formulas['thursday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format)#
        # self.formatted_cell(ws, start_row + 17, col_offset + 14, sales_formulas['friday'], font = self.font_bold, border = self.cell_border['bottom'], number_format=self.accounting_format)#
        # self.formatted_cell(ws, start_row + 17, col_offset + 16, sales_formulas['saturday'], font = self.font_green_bold, border = self.cell_border['bottomleft'], number_format=self.accounting_format) # These all rely on daily totals

        count_success_formulas = {day: '=' + ' + '.join([f'COUNTIF(C{self.cat_row_info[cat][0]}:C{self.cat_row_info[cat][1]}, "{date_strs[day]}")'for cat in self.cats_count_for_total]) for day in days}

        self.formatted_cell(ws, start_row + 3, col_offset + 10, count_success_formulas['monday']) # Actual count of monday jobs
        self.formatted_cell(ws, start_row + 3, col_offset + 11, count_success_formulas['tuesday']) # Actual count of tuesday jobs
        self.formatted_cell(ws, start_row + 3, col_offset + 12, count_success_formulas['wednesday']) # Actual count of wednesday jobs
        self.formatted_cell(ws, start_row + 3, col_offset + 13, count_success_formulas['thursday']) # Actual count of thursday jobs
        self.formatted_cell(ws, start_row + 3, col_offset + 14, count_success_formulas['friday']) # Actual count of friday jobs

        total_wk_count_success_formula = '=' + '+'.join([f'SUM({get_column_letter(col_offset + 10 + i)}{start_row + 3})' for i in range(5)])
        self.formatted_cell(ws, start_row + 3, col_offset + 15, total_wk_count_success_formula)

        self.formatted_cell(ws, start_row + 3, col_offset + 16, count_success_formulas['saturday']) # Actual count of saturday jobs
        self.formatted_cell(ws, start_row + 3, col_offset + 17, count_success_formulas['sunday']) # Actual count of sunday jobs

        count_unsuccess_formulas = {day: '=' + ' + '.join([f'COUNTIF(C{self.cat_row_info[cat][0]}:C{self.cat_row_info[cat][1]}, "{date_strs[day]}")'for cat in self.cats_count_for_unsuccessful]) for day in days}

        # Unsuccessful counts
        self.formatted_cell(ws, start_row + 4, col_offset + 10, count_unsuccess_formulas['monday'])
        self.formatted_cell(ws, start_row + 4, col_offset + 11, count_unsuccess_formulas['tuesday'])
        self.formatted_cell(ws, start_row + 4, col_offset + 12, count_unsuccess_formulas['wednesday'])
        self.formatted_cell(ws, start_row + 4, col_offset + 13, count_unsuccess_formulas['thursday'])
        self.formatted_cell(ws, start_row + 4, col_offset + 14, count_unsuccess_formulas['friday'])

        total_wk_count_unsuccess_formula = '=' + '+'.join([f'SUM({get_column_letter(col_offset + 10 + i)}{start_row + 4})' for i in range(5)])
        self.formatted_cell(ws, start_row + 4, col_offset + 15, total_wk_count_unsuccess_formula)

        self.formatted_cell(ws, start_row + 4, col_offset + 16, count_unsuccess_formulas['saturday'])
        self.formatted_cell(ws, start_row + 4, col_offset + 17, count_unsuccess_formulas['sunday'])

        curr_col = col_offset + 10
        for i in range(8):
            curr_col_letter = get_column_letter(curr_col + i)
            # success rate
            self.formatted_cell(ws, start_row + 5, curr_col + i, f'={curr_col_letter}{start_row + 5-2}/({curr_col_letter}{start_row + 5-2}+{curr_col_letter}{start_row + 5-1})', number_format=self.percentage_format)
            # avg sale
            self.formatted_cell(ws, start_row + 6, curr_col + i, f'={curr_col_letter}{start_row + 5-3}/{curr_col_letter}{start_row + 5-2}', number_format=self.accounting_format)

        return

    def job_section_title(self, row):
        ws = self.curr_worksheet
        col_offset = self.col_offset

        ws.cell(row, col_offset + 1, 'JOB DETAILS').border = self.cell_border['topleft']
        ws.cell(row, col_offset + 2).border = self.cell_border['top']
        ws.cell(row, col_offset + 3).border = self.cell_border['top']
        ws.cell(row, col_offset + 4).border = self.cell_border['top']
        ws.cell(row, col_offset + 5, 'JOB AMOUNT').border = self.cell_border['topleft']
        ws.cell(row, col_offset + 6).border = self.cell_border['top']
        ws.cell(row, col_offset + 7).border = self.cell_border['top']
        ws.cell(row, col_offset + 8).border = self.cell_border['top']
        ws.cell(row, col_offset + 9).border = self.cell_border['top']
        ws.cell(row, col_offset + 10).border = self.cell_border['topleft']
        ws.cell(row, col_offset + 11, 'PHOTOS').border = self.cell_border['topleft']
        ws.cell(row, col_offset + 12).border = self.cell_border['top']
        ws.cell(row, col_offset + 13).border = self.cell_border['top']
        ws.cell(row, col_offset + 14, 'QUOTE').border = self.cell_border['topleft']
        ws.cell(row, col_offset + 15).border = self.cell_border['top']
        ws.cell(row, col_offset + 16).border = self.cell_border['top']
        ws.cell(row, col_offset + 17, 'INVOICE').border = self.cell_border['topleft']
        ws.cell(row, col_offset + 18).border = self.cell_border['top']
        ws.cell(row, col_offset + 19).border = self.cell_border['top']
        ws.cell(row, col_offset + 20).border = self.cell_border['top']
        ws.cell(row, col_offset + 21, 'NOTES').border = self.cell_border['topright']

        row += 1
        self.curr_row = row

        ws.cell(row, col_offset + 1, 'JOB STATUS').border = self.cell_border['bottomleft']
        ws.cell(row, col_offset + 2, 'DATE').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 3, 'JOB #').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 4, 'SUBURB').border = self.cell_border['bottomright']
        ws.cell(row, col_offset + 5, 'AMOUNT EXC. GST').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 6, 'MATERIALS').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 7, 'MERCHANT FEES').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 8, 'NET PROFIT').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 9, 'PAID').border = self.cell_border['bottomright']
        ws.cell(row, col_offset + 10, 'DOC CHECK DONE').border = self.cell_border['bottomright']
        ws.cell(row, col_offset + 11, 'BEFORE').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 12, 'AFTER').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 13, 'RECEIPT').border = self.cell_border['bottomright']
        ws.cell(row, col_offset + 14, 'DESCRIPTION').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 15, 'SIGNED').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 16, 'EMAILED').border = self.cell_border['bottomright']
        ws.cell(row, col_offset + 17, 'DESCRIPTION').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 18, 'SIGNED').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 19, 'EMAILED').border = self.cell_border['bottomright']
        ws.cell(row, col_offset + 20, '5 Star Review').border = self.cell_border['bottom']
        ws.cell(row, col_offset + 21).border = self.cell_border['bottomright']
        ws.cell(row, col_offset + 22, 'EFTPOS').border = self.cell_border['bottomtop']
        ws.cell(row, col_offset + 23, 'CASH').border = self.cell_border['bottomtop']
        ws.cell(row, col_offset + 24, 'Payment Plan').border = self.cell_border['bottomtopright']
        
        self.curr_row += 1
        return
    
    def put_job_row(self, job: dict, row: int, cat_font: Font):
        ws = self.curr_worksheet
        col_offset = self.col_offset

        if job['first_appt_start_str'] != self.curr_date:
            job_border = self.cell_border['topthin']
        else:
            job_border = None
        self.curr_date = job['first_appt_start_str']

        if job['complaint_tag_present']:
            self.formatted_cell(ws, row, col_offset, 'COMPLAINT', font=cat_font)#, border = job_border)
        self.formatted_cell(ws, row, col_offset + 1, self.job_count, font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 2, job['first_appt_start_str'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 3, int(job['num']), font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 4, job['suburb'], font=cat_font, border = job_border)
        if not job['unsuccessful']:
            self.formatted_cell(ws, row, col_offset + 5, job['inv_subtotal'], font=cat_font, border = job_border, number_format=self.accounting_format)
            self.formatted_cell(ws, row, col_offset + 6, job['inv_subtotal']*0.2, font=cat_font, border = job_border, number_format=self.accounting_format)
            self.formatted_cell(ws, row, col_offset + 6, font=cat_font, border = job_border).comment = Comment(job['summary'], "automation")
        else:
            self.formatted_cell(ws, row, col_offset + 5, job['open_est_subtotal'], font=cat_font, border = job_border, number_format=self.accounting_format)
            self.formatted_cell(ws, row, col_offset + 6, "", font=cat_font, border = job_border)
        # 7
        self.formatted_cell(ws, row, col_offset + 7, "", font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 8, f"={get_column_letter(col_offset + 5)}{row}-{get_column_letter(col_offset + 6)}{row}-{get_column_letter(col_offset + 7)}{row}", font=cat_font, border = job_border, number_format=self.accounting_format)
        self.formatted_cell(ws, row, col_offset + 9, job['payment_types'], font=cat_font, border = job_border)
        # 10 TODO: all doc checks complete
        doc_check_complete_col = f'=IF(OR({", ".join([f"{get_column_letter(col_offset + i)}{row}=0" for i in range(11,20)])}), "N","Y")'
        self.formatted_cell(ws, row, col_offset + 10, doc_check_complete_col, font=cat_font, border = job_border).alignment = self.align_center
        self.formatted_cell(ws, row, col_offset + 11, job['Before Photo'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 12, job['After Photo'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 13, job['Receipt Photo'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 14, job['Quote Description'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 15, job['Quote Signed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 16, job['Quote Emailed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 17, job['Invoice Description'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 18, job['Invoice Signed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 19, job['Invoice Emailed'], font=cat_font, border = job_border)
        self.formatted_cell(ws, row, col_offset + 20, job['5 Star Review'], font=cat_font, border = job_border)

        self.formatted_cell(ws, row, col_offset + 25, f"=ROUND(F{row}*1.1,2)", font=self.font_purple)
        self.formatted_cell(ws, row, col_offset + 26, f"=ROUND(Z{row} - W{row} - X{row} - Y{row},2)", font=self.font_purple)
        
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
                self.formatted_cell(ws, row, col_offset + 22, f"={'+'.join(p_types['EF'] + p_types['Cr'])}", font=cat_font, number_format=self.accounting_format)
            if p_types['Ca']:
                self.formatted_cell(ws, row, col_offset + 23, f"={'+'.join(p_types['Ca'])}", font=cat_font, number_format=self.accounting_format)
        return
    
    def category_totals_row(self, cat: str, amt_col: int, materials_col: int, merchantf_col: int, profit_col: int):
        # totals row
        ws = self.curr_worksheet
        col_offset = self.col_offset
        curr_row = self.curr_row
        cat_row_start, cat_row_end = self.cat_row_info[cat]

        amt_letter = get_column_letter(col_offset + amt_col)
        materials_letter = get_column_letter(col_offset + materials_col)
        merchantf_letter = get_column_letter(col_offset + merchantf_col)
        profit_letter = get_column_letter(col_offset + profit_col)
        ws.cell(curr_row, col_offset + amt_col, f"=SUM({amt_letter}{cat_row_start}:{amt_letter}{cat_row_end})").number_format = self.accounting_format
        ws.cell(curr_row, col_offset + materials_col, f"=SUM({materials_letter}{cat_row_start}:{materials_letter}{cat_row_end})").number_format = self.accounting_format
        ws.cell(curr_row, col_offset + merchantf_col, f"=SUM({merchantf_letter}{cat_row_start}:{merchantf_letter}{cat_row_end})").number_format = self.accounting_format
        ws.cell(curr_row, col_offset + profit_col, f"=SUM({profit_letter}{cat_row_start}:{profit_letter}{cat_row_end})").number_format = self.accounting_format
        ws.cell(curr_row, col_offset + amt_col).border = self.border_double_bottom
        ws.cell(curr_row, col_offset + materials_col).border = self.border_double_bottom
        ws.cell(curr_row, col_offset + merchantf_col).border = self.border_double_bottom
        ws.cell(curr_row, col_offset + profit_col).border = self.border_double_bottom

        if cat == 'wk_complete_paid':
            for i in range(10):
                curr_col = col_offset + 11 + i
                self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
                self.formatted_cell(ws, curr_row + 1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            self.curr_row += 1

        if self.bottom_row < self.curr_row:
            self.bottom_row = self.curr_row
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # curr_col += 1
            # self.formatted_cell(ws, curr_row, curr_col, f"=COUNT({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")


            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
            # self.formatted_cell(ws, curr_row+1, curr_col, f"=SUM({get_column_letter(curr_col)}{cat_row_start}:{get_column_letter(curr_col)}{cat_row_end})")
        return
    
    def put_job_category(self, cat: str, cat_text: str, job_cats:dict, start_row: int):
        ws = self.curr_worksheet
        col_offset = self.col_offset
        cat_font = None
        
        if cat in ['wk_complete_unpaid', 'wk_wo', 'wkend_complete_unpaid', 'wkend_wo']:
            cat_font = self.font_green
        self.formatted_cell(ws, start_row, col_offset + 1,cat_text)#, font=cat_font)
        cat_row = start_row
        cat_row += 1

        if cat == 'prev':
            for row in ws[f'B{cat_row-1}:X{cat_row+7}']: # TODO: Make these column letters dynamic if possible?
                for cell in row:
                    cell.fill = self.yellow_fill
            for i in range(1,9):
                ws.cell(cat_row, col_offset + 1, i)
                cat_row += 1

        else:
            jobs = job_cats.get(cat, [])
            if not jobs:
                cat_row += 1
                # self.curr_row = cat_row
            self.job_count = 1
            for job in jobs:
                self.put_job_row(job, cat_row, cat_font)
                cat_row += 1
                self.job_count += 1
        
        self.curr_row = cat_row # Don't need to mess around with curr_row this much, just making sure it works for now.
        # if self.bottom_row < self.curr_row:
        #     self.bottom_row = self.curr_row
        
        # add 5 rows for any extras to add
        self.curr_row += 5
        if self.bottom_row < self.curr_row:
            self.bottom_row = self.curr_row

        # Fill in info 
        self.cat_row_info[cat] = (start_row + 1, self.curr_row - 1)

        self.category_totals_row(cat, amt_col=5, materials_col=6, merchantf_col=7, profit_col=8)
        self.curr_row += 2
        if self.bottom_row < self.curr_row:
            self.bottom_row = self.curr_row
        return

    def extra_formatting(self):
        ws = self.curr_worksheet

        # ----------------------------------------------------------------------------------------
        # Doc check formatting 0s to red
        dc_start_col_letter = get_column_letter(self.col_offset + 10)
        dc_end_col_letter = get_column_letter(self.col_offset + 20)
        for cat in self.cat_row_info.keys():
            cat_row_start, cat_row_end = self.cat_row_info[cat]
            dc_range = f'{dc_start_col_letter}{cat_row_start}:{dc_end_col_letter}{cat_row_end}'
            ws.conditional_formatting.add(
                dc_range, 
                CellIsRule(operator="equal", formula=["0"], font=self.font_red)
            )
        # ----------------------------------------------------------------------------------------
        # Highlighting red for no doc check
        ws.conditional_formatting.add(
            f"I{self.cat_row_info['wk_complete_paid'][0]}:I{self.cat_row_info['ah_wo'][1]}", 
            FormulaRule(formula=[f'$K{self.cat_row_info["wk_complete_paid"][0]}="N"'], fill=PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"))
        )
        # ----------------------------------------------------------------------------------------

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
        
        # ========================== BOTTOM SECTION - JOBS ==========================
        job_start_row = 50
        # self.curr_row = job_start_row
        self.job_section_title(job_start_row)

        for cat, cat_text in self.CATEGORY_ORDER.items():
            self.put_job_category(cat, cat_text, job_cats, self.curr_row)
        
        self.extra_formatting()

        # ========================= TOP SECTION - SUMMARIES =========================
        # Will need to call some of these after doing jobs because they will need to know cat start and ends etc.
        self.title_row(start_row=1)
        self.job_count_box(start_row=3)
        self.profit_target_box(start_row=3)
        self.payout_box(start_row=7)
        self.doc_check_count_box(start_row=2)
        if self.timeframe == 'weekly':
            self.day_summaries_weekly(start_row=8)
        elif self.timeframe == 'monthly':
            self.day_summaries_monthly(start_row=8)
        else:
            raise ValueError("Timeframe must be 'weekly' or 'monthly'")

    def build_workbook(self):
        # Final function that actually builds everything
        wb = Workbook()

        for tech in sorted(self.jobs_by_tech.keys()):
            self.build_sheet(wb, tech)

        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)
        return bio.getvalue()
