from openpyxl import Workbook   #modulo per manipolare file excel
from openpyxl.chart import (    #moduli per creare grafici su excel
    LineChart,
    Reference,
)
import os
import serial

altro='S'



import re

txt='start time: 10:32:50.20'

re1='(start)'	# Word 1
re2='(\\s+)'	# White Space 1
re3='(time)'	# Word 2

rg = re.compile(re1+re2+re3,re.IGNORECASE|re.DOTALL)
m = rg.search(txt)
if m:
    word1=m.group(1)
    ws1=m.group(2)
    word2=m.group(3)
    print "("+word1+")"+"("+ws1+")"+"("+word2+")"+"\n"



print('~=~=Spettrofotometro=~=~')
print('')
porta = serial.Serial('COM7')     #apri porta seriale
wb = Workbook()                 #crea un file excel
wb.remove(wb.active)            #rimuove il foglio di default
while altro=='S':
    nome = input('Nome del campione: ')         #inserire il nome del campione da analizzare
    ws = wb.create_sheet(nome)
    dati = []                                   #crea l'input
    print('\nAttendo i dati...')
    while True:
        carattere = porta.read()                #legge 1 byte
        if dati==bytes('\r','utf-8'):
            stringa = ''.join(dati)             #converte la lista in stringa
            if stringa=='WL    AU':
                scan=True                       #la scansione è iniziata!
            if stringa=='start time'
            stringa=[]
        elif dati==bytes('\n','utf-8'):
            pass
        else:
            stringa.append(str(dati, 'utf-8'))
    for a in range(1,10):
        for b in range(1,10):
            ws.cell(row=a, column=b, value=a+b)
    grafico = LineChart()
    grafico.title = "Curva assorbanza di "+nome
    grafico.style = 13
    grafico.y_axis.title = 'Assorbanza'
    grafico.x_axis.title = 'Lunghezza d\'onda'
    altro = 'C'
    while altro!='S' and altro!='N':
        altro = input('Inserire un altro campione? (S/N)')

wb.save(os.environ['UserProfile']+'\\Desktop\\dati.xlsx')
