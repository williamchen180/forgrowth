import sys
import traceback
import json
import datetime
import time
import random

import openpyxl
import requests

my_work_book = 'growth.xlsx'

wb_obj = openpyxl.load_workbook(my_work_book)

print(wb_obj.sheetnames)


def calculate_percentage(buy, sell):
    return (sell - buy) / buy


def convert_to_str(input):
    if type(input) is str:
        return input.rstrip(' ').lstrip(' ').rstrip(' ')

    if type(input) is int:
        return str(input)

    return input


def load_cell(s, r, c):
    x = s.cell(row=r, column=c)
    x.value = convert_to_str(x.value)
    return x


def load_value(s, r, c):
    return load_cell(s, r, c).value


def set_cell(s, r, c, value):
    cell = s.cell(row=r, column=c)
    cell.value = value


def column_idx(char):
    return ord(char.upper()) - ord('A') + 1


iteration_times = 1000
number_candidates = 2


history_gain = []

sheet = wb_obj['分析']
history_trade = []


column = ord('U') - ord('A') + 1
skip = []
for i in range(62, sheet.max_row + 1):

    if i in skip:
        continue

    count = 0
    if load_cell(sheet, i, 3).value == '投信買籌多':

        date = load_value(sheet, i, column_idx('B'))

        date = load_value(sheet, i, column_idx('B'))
        symbol = load_value(sheet, i, column_idx('D'))
        name = load_value(sheet, i, column_idx('E'))
        base_price = load_value(sheet, i, column_idx('K'))
        percentage = load_value(sheet, i, column_idx('O'))  # 4 day
        rand_idx = load_value(sheet, i, column_idx('U'))

        history_trade.append((date, symbol, name, base_price, percentage, rand_idx))

sheet = wb_obj['模擬交易']

columns = 3
for i in range(0, iteration_times):

    total_gain = 0
    r = random.sample(range(10), 10)
    candidates = [str(x) for x in r[0:number_candidates]]
    print('candidate', candidates)

    rows = 3
    for x in history_trade:
        print(x)
        (date, symbol, name, base_price, percentage, rand_idx) = x
        if percentage is None:
            continue
        if rand_idx not in candidates:
            continue

        gain = float(base_price)*float(percentage)
        total_gain = total_gain + gain

        #print(date, symbol, name, base_price, gain)

        set_cell(sheet, rows, columns+0, date)
        set_cell(sheet, rows, columns+1, symbol)
        set_cell(sheet, rows, columns+2, name)
        set_cell(sheet, rows, columns+3, base_price)
        set_cell(sheet, rows, columns+4, gain)

        rows = rows + 1

    set_cell(sheet, 2, columns+3, total_gain)
    #print('Total: %.2f' % total_gain)

    columns = columns + 6

    set_cell(sheet, i+2, 1, total_gain)
    history_gain.append(total_gain)

wb_obj.save(my_work_book)

for x in history_gain:
    print(x)
