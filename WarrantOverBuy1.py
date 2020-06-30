
import csv
import openpyxl
from datetime import datetime
from libexcel import *
my_work_book = 'WarrantOverBuy.xlsx'

wb_obj = openpyxl.load_workbook(my_work_book)

print(wb_obj.sheetnames)

#sheet = wb_obj['買超購']
#sheet = wb_obj['買超售']
#sheet = wb_obj['賣超購']
sheet = wb_obj['賣超售']

minute_mark = [5, 15, 30, 45]

bandit1 = 0.001425
bandit2 = 0.0015

for i in range(13, sheet.max_row + 1):
    date = load_value(sheet, i, column_number('A'))
    symbol = load_value(sheet, i, column_number('B'))

    print(date, symbol)

    if symbol is None or date is None:
        continue

    csv_file = f'KLINE/K_{symbol}.csv'

    with open(csv_file) as f:
        rows = csv.reader(f)

        count = 0
        column = column_number('L')
        day_found = False
        base_price = 0
        for r in rows:
            d = datetime.strptime(r[0], '%Y/%m/%d %H:%M')

            if day_found is True and (d.year, d.month, d.day) != (date.year, date.month, date.day):
                print('finished')
                break

            if day_found is not True and (d.year, d.month, d.day) == (date.year, date.month, date.day):
                print('day found')
                day_found = True


            if day_found is True:

                if d.hour == 9 and d.minute == 5:
                    base_price = float(r[1])
                    set_cell(sheet, i, column_number('K'), 0)

                if base_price != 0:
                    current_profit_rate = (base_price * (1.0 - bandit1) - float(r[1]) * (1.0 + bandit2 + bandit1)) / base_price
                    if current_profit_rate > 0.01:
                        print("1% achieved: ", current_profit_rate)
                        ratio = load_value(sheet, i, column_number('I'))
                        if ratio is None:
                            ratio = 0.0
                        else:
                            ratio = float(ratio)
                        if ratio < current_profit_rate:
                            set_cell(sheet, i, column_number('I'), current_profit_rate)
                            pass

                if d.minute in minute_mark:
                    set_cell(sheet, i, column, r[4])
                    count = count + 1
                    column = column + 1


wb_obj.save(my_work_book)

