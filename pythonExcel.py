import xlrd
loc = ("OktaTest.xls")
wb=xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(0)
rows=sheet.nrows
i=0
for i in range(rows):
    print(sheet.cell_value(i,0))
