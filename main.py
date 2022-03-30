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
thead = ['Дата', 'Значение курса промежуточного клиринга ', 'Время', 'Значение курса основного клиринга ', 'Время']

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

def join(usd, eur):
    data = {}
    el = ['-', '-', '-', '-', '-']
    for u in usd:
        data[u[0]] = u + el
    for e in eur:
        if list(data.keys()).count(e[0]) == 0:
            data[e[0]] = el + e[1:]
        else:
            data[e[0]] = data[e[0]][:5] + e
    return data

def createFile(data):
    file = xl.Workbook()
    sheet = file.active
    sheet['J1'] = 'Изменение'
    sheet.cell(row=1, column=9).alignment = Alignment(horizontal='center')
    col = 1
    for c in range(len(currs)):
        for t in range(len(thead)):
            h = thead[t]
            if h.find('курса') != -1:
                h += currs[list(currs.keys())[c]]
            sheet.cell(row=1, column=col, value=h)
            sheet.cell(row=1, column=col).alignment = Alignment(horizontal='center')
            col += 1
    row = 2

    for k in data.keys():
        d = data[k]
        for i in range(9):
            val = d[i]
            if val.count(':') == 0 and val.count('.') == 1:
                val = float(val)
            sheet.cell(row=row, column=i + 1, value=val)
            sheet.cell(row=row, column=i + 1).alignment = Alignment(horizontal='center')

        val = '-'
        if d[3] != '-' and d[-2] != '-':
            val = float(d[-2]) / float(d[3])
        sheet.cell(row=row, column=10, value=val)
        sheet.cell(row=row, column=10).alignment = Alignment(horizontal='center')
        row += 1

    for l in ['B', 'D', 'G',  'I']:
        sheet.column_dimensions[l].width = 46
    file.save('data.xlsx')

res['usd'] = BeautifulSoup(requests.get(moex_url + currs['usd']).text, 'lxml')
res['eur'] = BeautifulSoup(requests.get(moex_url + currs['eur']).text, 'lxml')
res['usd'] = res['usd'].find('table', class_='tablels').find_all('tr')[2:]
res['eur'] = res['eur'].find('table', class_='tablels').find_all('tr')[2:]
usd = makeDict(res['usd'])
eur = makeDict(res['eur'])
data = join(usd, eur)
createFile(data)
