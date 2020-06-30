
import csv
import openpyxl
from datetime import datetime
from libexcel import *
my_work_book = 'WarrantOverBuy.xlsx'

wb_obj = openpyxl.load_workbook(my_work_book)

print(wb_obj.sheetnames)

sheet = wb_obj['Source']


for i in range(2, sheet.max_row + 1):
    date = load_value(sheet, i, column_number('A'))
    symbol = load_value(sheet, i, column_number('B'))

    print(date, symbol)

    csv_file = f'KLINE/K_{symbol}.csv'

    with open(csv_file) as f:
        rows = csv.reader(f)

        for r in rows:
            d = datetime.strptime(r[0], '%Y/%m/%d %H:%M')

            if (d.year, d.month, d.day) == (date.year, date.month, date.day):
                print('day found')

                for

