import json
from xlrd import open_workbook
import xlsxwriter


column_ident = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
start_at = 3
wb = open_workbook('sheeds.xlsx')
# print number of sheets
print wb.nsheets
 
# print sheet names
print wb.sheet_names()
data = []
for s in wb.sheets():
    #print s.__dict__
    #print s.xlrd
    break
    #print 'Sheet:',s.name
    values = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value  = (s.cell(row,col).value)
            try : value = str(int(value))
            except : pass
            col_value.append(value)
        values.append(col_value)
    data.append(values)
#print json.dumps(data)
workbook = xlsxwriter.Workbook('sheeds_out.xlsx')
worksheet = workbook.add_worksheet('baocao')
header_fortmat = workbook.add_format({'bold': True, 'font_color': 'red'})
for i, l in enumerate(values):
    i = i + start_at
    for j, col in enumerate(l):
        worksheet.write(i, j, col, header_fortmat)
#        print "--------------------"
#        print i
#        print j
#        print l
#        print col
#        print "--------------------"
workbook.close()

