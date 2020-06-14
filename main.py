import openpyxl
import excel_utils
import logging
from log_utils import logd, logi, logw, loge, set_log_level


set_log_level(logging.INFO)

# filename = 'test/random.xlsx'
filename = 'E:\gp\\6_2.xlsx'
file_69 = 'E:\gp\origin\\6.9.xlsx'
file_610 = 'E:\gp\origin\\6.10.xlsx'
file_611 = 'E:\gp\origin\\6.11.xlsx'
file_612 = 'E:\gp\origin\\6.12.xlsx'

excel_utils.create_excel(filename)

def process_excel(finename, sheetname, filetemp):
    wb = openpyxl.load_workbook(finename)
    # sheet = wb.active
    sheet = wb.create_sheet(sheetname)
    logi('start file %s------------------------', filetemp)
    wb_temp = openpyxl.load_workbook(filetemp)
    sheet_temp = wb_temp.active

    sheet_temp = excel_utils.find_and_delete_key(sheet_temp, 'SH688', 1)
    sheet_temp = excel_utils.find_and_delete_key(sheet_temp, 'ST', 1)
    sheet_temp = excel_utils.delete_range(sheet_temp, 9)
    for row in range(2, sheet_temp.max_row + 1):
        logi('%s, %s', sheet_temp.cell(row, 1).value, sheet_temp.cell(row, 2).value)
    for row in range(1, sheet_temp.max_row+1):
        for cell in sheet_temp[row]:
            sheet.cell(row, cell.column).value = cell.value
# wb.create_sheet('new', 2)
    wb.save(filename)

process_excel(filename, '6.9', file_69)
process_excel(filename, '6.10', file_610)
process_excel(filename, '6.11', file_611)
process_excel(filename, '6.12', file_612)