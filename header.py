import json
from xlrd import open_workbook
import xlsxwriter
import os.path

column_ident = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
start_at = 3
tencent = [
	["Tencent's tool", "", ""],
        ["Bill", "", ""],
        ["Date", "Update User", "Cost ($)"]
]

vng = [
        ["VNG's KPI tool", "", ""],
        ["","",""],
        ["Date", "A1", "Tencent Update user < VNG A1"]
]


def get_max_col(data_list):
    _max = 0
    for row in data_list:
        if _max < len(row):
            _max = len(row)
    return _max

def xls_write(data_list=None, name="abc", sheet="sheet1", row_start_at=0, col_start_at=0):
    try:
        if os.path.isfile(fname):
            pass 
        workbook = xlsxwriter.Workbook(name)
        worksheet = workbook.add_worksheet(sheet)
        header_fortmat = workbook.add_format({'bold': True, 'font_color': 'red'})
    except Exception as e:
        raise e
        return 1
    try:
        for i, l in enumerate(data_list):
            i = i + row_start_at
            for j, col in enumerate(l):
                j = j + col_start_at
                worksheet.write(i, j, col, header_fortmat)
                print "--------------------"
                print i
                print j
                print l
                print col
                print "--------------------"
        workbook.close()
    except Exception as e:
        raise e
        return 1
    return 0

if __name__ == "__main__":
    xls_write(data_list=tencent, name="header.xls", sheet="baocao", row_start_at=1, col_start_at=0)
    start_at = get_max_col(vng) + 2
    xls_write(data_list=vng, name="header.xls", sheet="baocao", row_start_at=1, col_start_at=start_at)
