from openpyxl import Workbook   #modulo per manipolare file excel
from openpyxl.chart import (    #moduli per creare grafici su excel
    LineChart,
    Reference,
)
import os
import serial

print('~=~=Spettrofotometro=~=~')
print('Ricerca porta seriale...')
for i in range(1,255):
    try:
        ser = serial.Serial('COM7')  # open serial port
ser.write(bytes(ster, 'utf-8'))
wb = Workbook()
ws= wb.active
for a in range(1,10):
    for b in range(1,10):
        ws.cell(row=a, column=b, value=a+b)
wb.save(os.environ['UserProfile']+'\\dati.xlsx')
