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
current_date = ''
for row in sheet.rows:
    values = [x.value for x in row]

    # 把 title 寫到其他表頭
    if title is True:
        title = False
        column = 1
        for x in values:
            set_cell(sheet1, 1, column, x)
            set_cell(sheet2, 1, column, x)
            set_cell(sheet3, 1, column, x)
            column = column + 1
        continue

    date = values[0]
    policy = values[1]

    print(f'date: {date}, policy: {policy}, type(policy): {type(policy)}')

    if policy == '1' or policy == 1:
        if current_date != date:
            current_date = date
            data1.append([])
            print(f'new day: {date}')

        data1[-1].append(values)

#
# data1:
#
# [ [ [D1S1], [D1S2], [D1S3]... ], [ [D2S1], [D2S2], [D2S3]..], ... ]
#


print(f'days: {len(data1)}')

for this_day in data1:
    print(f'{len(this_day)} stocks of today')


array1 = numpy.array(data1)

print(array1)
print(array1.shape)




#wb_obj.save(my_work_book)




