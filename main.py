from openpyxl import Workbook   #modulo per manipolare file excel
from openpyxl.chart import (    #moduli per creare grafici su excel
    LineChart,
    Reference,
)

import serial

ster='Banana'
print(ster)
ser = serial.Serial('COM7')  # open serial port
ser.write(bytes(ster, 'utf-8'))
while True:
    btr = ser.inWaiting()
    data = ser.    print(data)
#wb = Workbook()
#ws= wb.active
#for a in range(1,10):
#    for b in range(1,10):
#        ws.cell(row=a, column=b, value=a+b)
#wb.save('dati.xlsx')
