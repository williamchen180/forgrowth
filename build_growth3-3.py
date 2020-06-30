#
# 真正做分析
#
# 目前標頭為
#
# 日期	策略	代號	名稱	成交價	  漲跌    幅(%)  	成交量	  月線   斜率	  上通   斜率	  下通   斜率	帶寬	位階	主力連買	外資連買	投信連買	主1	主5	主10	外資%	投信%	自避%	當沖%	融資%	融券%	券資比%	乖離年線%	融資維持率	當沖損益(千)
#
# 輸出在 [輸出] 的sheet
#
#

import os
import sys
import traceback
import json
import requests
import datetime
import openpyxl
import pickle
from libexcel import *
import numpy


def handle_data1(sheet, data, tops=10):
    array = numpy.array(data)

    item_written = 0

    # sort by 投信
    array = array[array[:, column_number('T')-1].argsort()][::-1]

    #print(sheet._current_row)

    for stock in array:

        # 交易量小於300去除
        if stock[column_number('G')-1] < 300:
            continue

        if item_written >= tops:
            continue

        item_written = item_written + 1

        sheet._current_row = sheet._current_row + 1
        sheet._current_column = 1
        for value in stock:
            #print(f'set_cell {sheet._current_row} {sheet._current_column} {value}')
            set_cell(sheet, sheet._current_row , sheet._current_column, value)
            sheet._current_column = sheet._current_column + 1


def handle_data2(sheet, data, tops=9999):
    array = numpy.array(data)

    item_written = 0

    # sort by 投信
    array = array[array[:, column_number('T')-1].argsort()][::-1]

    #print(sheet._current_row)

    for stock in array:

        # 交易量小於300去除
        if stock[column_number('G')-1] < 300:
            continue

        if item_written >= tops:
            continue

        item_written = item_written + 1

        sheet._current_row = sheet._current_row + 1
        sheet._current_column = 1
        for value in stock:
            #print(f'set_cell {sheet._current_row} {sheet._current_column} {value}')
            set_cell(sheet, sheet._current_row , sheet._current_column, value)
            sheet._current_column = sheet._current_column + 1


my_work_book = 'growth3.xlsx'
wb_obj = openpyxl.load_workbook(my_work_book)

print(wb_obj.sheetnames)
sheet = wb_obj['分析']

sheet1 = wb_obj['策略1']
sheet2 = wb_obj['策略2']
sheet3 = wb_obj['策略3']

data1 = []
data2 = []
data3 = []

title = True
# 日期	策略	代號	名稱	成交價	  漲跌    幅(%)  	成交量	  月線   斜率	  上通   斜率	  下通   斜率	帶寬	位階	主力連買	外資連買	投信連買	主1	主5	主10	外資%	投信%	自避%	當沖%	融資%	融券%	券資比%	乖離年線%	融資維持率	當沖損益(千)		開盤	1 day	2 day	3 day	4 day	5 day	10 day	15 day	20 day	30 days
current_date1 = ''
current_date2 = ''
current_date3 = ''
for row in sheet.rows:
    values = [x.value for x in row]

    # 把 title 寫到其他表頭
    if title is True:
        title = False
        column = 1
        for x in values:
            set_cell(sheet1, 12, column, x)
            set_cell(sheet2, 12, column, x)
            set_cell(sheet3, 12, column, x)
            column = column + 1
        sheet1._current_row = 12
        sheet2._current_row = 12
        sheet3._current_row = 12
        continue

    date = values[0]
    policy = values[1]

    if policy == '1' or policy == 1:
        if current_date1 != date:
            current_date1 = date
            if len(data1) != 0:
                handle_data1(sheet1, data1)
            data1 = []
            print(f'new day: {date}')
        data1.append(values)

    if policy == '2' or policy == 2:
        if current_date2 != date:
            current_date2 = date
            if len(data2) != 0:
                handle_data2(sheet2, data2)
            data2 = []
            print(f'new day: {date}')
        data2.append(values)

    if policy == '3' or policy == 3:
        if current_date3 != date:
            current_date3 = date
            if len(data3) != 0:
                handle_data2(sheet3, data3)
            data3 = []
            print(f'new day: {date}')
        data3.append(values)

if len(data1) != 0:
    handle_data1(sheet1, data1)

if len(data2) != 0:
    handle_data2(sheet2, data2)

if len(data3) != 0:
    handle_data2(sheet3, data3)

wb_obj.save(my_work_book)



