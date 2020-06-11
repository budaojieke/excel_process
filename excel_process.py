import xlrd
import xlwt

def read_excel(bookname,  sheetname):
    # 打开Excel文件
        data = xlrd.open_workbook(bookname)
        sheet = data.sheet_by_name(sheetname)
        dic = {}

        if 1:
        # 把第一列作为字典的键，一行数据保存为列表，作为字典的值。列表中也包含第一列的值哦
            for i in range(sheet.nrows):
                lis = []
                for j in range(sheet.ncols):
                    lis.append(sheet.cell(i, j).value)
                    # print(i, j)
                dic[sheet.cell(i, 0).value] = lis
        else:
            for i in range(sheet.ncols):
                lis = []
                for j in range(sheet.nrows):
                    lis.append(sheet.cell(j, i).value)
                    # print(i, j)
                    # print(lis)
                dic[sheet.cell(0, i).value] = lis
            # print(dic)
        return dic


def find_and_delete(dic, string):
    delete_data = []
    for key in dic.keys():
        value = dic.get(key)
        r1 = string in value
        if r1:
            delete_data.append(key)
            # print(list(dic.keys())[list(dic.values()).index(string)])
    print(delete_data)
    for k in delete_data:
        dic.pop(k)
    return dic


def write_back(dic, filename):
    # 这个参数意思，我也不太明白，反正这么写就对了，字面意思理解即可
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('test', cell_overwrite_ok=True)

    m = 0

    for i in dic.keys():
        # 这个地方是从第二行开始写起了。。。
        m += 1
        n = 0
        for j in dic[i]:
            sheet.write(m, n, j)
            n += 1

    book.save(filename)
