import requests
from bs4 import BeautifulSoup
import openpyxl as xl
from openpyxl.styles import Alignment

moex_url = 'https://www.moex.com/ru/derivatives/currency-rate.aspx?currency='
currs = {
    'usd': 'USD_RUB',
    'eur': 'EUR_RUB'
}
res = {
    'usd': '',
    'eur': ''
}
thead = ['Значение курса промежуточного клиринга ', 'Время', 'Значение курса основного клиринга ', 'Время']

usd = []
eur = []

def makeDict(vals):
    dict = []
    for str in vals:
        a = []
        for s in str.find_all('td'):
            a.append((s.string + '').replace(',', '.'))
        dict.append(a)
    return dict

def createFile(usd, eur):
    file = xl.Workbook()
    sheet = file.active
    sheet['A1'] = 'Дата'
    sheet.cell(row=1, column=1).alignment = Alignment(horizontal='center')
    sheet['J1'] = 'Изменение'
    sheet.cell(row=1, column=9).alignment = Alignment(horizontal='center')
    col = 2
    for c in range(len(currs)):
        for t in range(len(thead)):
            h = thead[t]
            if col % 2 == 0:
                h += currs[list(currs.keys())[c]]
            sheet.cell(row=1, column=col, value=h)
            sheet.cell(row=1, column=col).alignment = Alignment(horizontal='center')
            col += 1
    row = 2
    for e in eur:
        for d in usd:
            r = []
            if e[0] == d[0]:
                r = d + e[1:]
                for i in range(9):
                    val = r[i]
                    if val.count(':') == 0 and val.count('.') == 1:
                        val = float(val)
                    sheet.cell(row=row, column=i+1, value=val)
                    sheet.cell(row=row, column=i+1).alignment = Alignment(horizontal='center')

                val = '-'
                if e[-2] != '-' and d[-2] != '-':
                    val = float(e[-2])/float(d[-2])
                sheet.cell(row=row, column=10, value=val)
                sheet.cell(row=row, column=10).alignment = Alignment(horizontal='center')
                row += 1
    for l in ['B', 'D', 'F',  'H']:
        sheet.column_dimensions[l].width = 46
    file.save('data.xlsx')

res['usd'] = BeautifulSoup(requests.get(moex_url + currs['usd']).text, 'lxml')
res['eur'] = BeautifulSoup(requests.get(moex_url + currs['eur']).text, 'lxml')
res['usd'] = res['usd'].find('table', class_='tablels').find_all('tr')[2:]
res['eur'] = res['eur'].find('table', class_='tablels').find_all('tr')[2:]
usd = makeDict(res['usd'])
eur = makeDict(res['eur'])
createFile(usd, eur)
