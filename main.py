from openpyxl import Workbook


wb = Workbook()
ws= wb.active
#ws['A3']=142
#ws['B7']=0.28
for a in range(1,10):
    for b in range(1,10):
        ws.cell(row=a, column=b, value=a+b)
wb.save('test.xlsx')
