import openpyxl


def find_and_delete_key(workSheet, key):
    # print(sheet.max_row)
    flag = 0
    for row in range(workSheet.max_row, 1, -1):
        for cell in workSheet[str(row)]:
            data = cell.value
            if key in str(data):
                flag = 1
                # print(cell, data, key)
                break
        if flag == 1:
            workSheet.delete_rows(row)
            flag = 0
    # workSheet.max_row = workSheet.max_row - i
    return workSheet




