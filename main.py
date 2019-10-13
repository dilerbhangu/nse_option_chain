import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import *
import requests
import json
import time
import sys


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('Credentials.json', scope)
client = gspread.authorize(creds)


url = 'https://beta.nseindia.com/api/option-chain-indices?symbol=NIFTY'

heading = ['OI', 'Chg in OI', 'IV', 'LTP', 'Strike Price', 'LTP', 'IV', 'Chg in OI', 'OI']
val2 = ['openInterest', 'changeinOpenInterest', 'impliedVolatility',
        'lastPrice', 'strikePrice', 'lastPrice', 'impliedVolatility', 'changeinOpenInterest', 'openInterest']
sh = client.open('Option Chain Data')
col_name = ['A2:A15', 'B2:B15', 'C2:C15', 'D2:D15',
            'E2:E15', 'F2:F15', 'G2:G15', 'H2:H15', 'I2:I15']
worksheet = None
lastPrice = 0


def fetch_oi():
    headers = {'User-Agent':
               'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'}
    r = requests.get(url, headers=headers).json()
    print(r)
    data = json.dumps(r, sort_keys=True)
    data = json.loads(data)
    return data


def process_oi(data):

    currentPrice = int(data['records']['underlyingValue'])
    list = data['filtered']['data']
    starting_index = 0

    for data in list:
        if(int(data['strikePrice']) > currentPrice):
            starting_index = starting_index-7
            break
        else:
            starting_index += 1

    return list, starting_index


def check_holiday():
    currentPrice = data['records']['underlyingValue']

    if lastPrice == currentPrice:
        priceCountCond += 1

    if priceCountCond == 5:
        sh.del_worksheet(worksheet)
        sys.exit()

    lastPrice = currentPrice


def update_excel_sheet(data, starting_index):

    if worksheet != None:
        val1 = 'CE'
        for col_index in range(0, 9):

            update_column(col_name[col_index], starting_index,
                          val1, val2[col_index], data)

            if val2[col_index] == 'strikePrice':
                val1 = 'PE'


def update_column(col_name, index, val1, val2, data):
    cell_list = worksheet.range(col_name)
    for cell in cell_list:
        cell.value = data[index][val1][val2]
        index += 1

    worksheet.update_cells(cell_list)


def create_excel_sheet():
    global worksheet
    worksheet = sh.add_worksheet(datetime.today().strftime("%d/%m/%Y"), 100, 20)
    index = 0
    cell_list = worksheet.range('A1:I1')
    for cell in cell_list:
        cell.value = heading[index]
        index += 1

    worksheet.update_cells(cell_list)


def wait_time(seconds):
    time.sleep(seconds)


def exit_cond():
    d = datetime.utcnow()
    hour = (d.hour+6) % 24
    minute = (d.minute+30) % 60

    if hour == 9:
        check_holiday()

    if hour == 15:
        if minute > 35:
            sys.exit()


def main():

    while True:
        data = fetch_oi()

        if worksheet == None:
            create_excel_sheet()

        data, starting_index = process_oi(data)
        update_excel_sheet(data, starting_index)
        wait_time(180)
        exit_cond()


if __name__ == "__main__":
    main()
