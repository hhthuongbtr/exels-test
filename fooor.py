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

for x, y in zip(tencent, vng):
    print x
    print y
