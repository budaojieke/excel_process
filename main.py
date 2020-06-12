import openpyxl
import excel_utils

filename = filename = 'test/random.xlsx'

wb = openpyxl.load_workbook(filename)
sheet = wb.active

new_sheet = excel_utils.find_and_delete_key(sheet, '0')


for row in range(2, new_sheet.max_row + 1):
    print(new_sheet.cell(row, 1).value)

# wb.create_sheet('new', 2)
wb.save(filename)

