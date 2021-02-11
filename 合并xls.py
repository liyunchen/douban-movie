# -*- coding: utf-8 -*-

import requests
from lxml import etree
import xlrd
import xlwt
import time
import random
from xlutils.copy import copy




fname="豆瓣"

datalist=[]
for i in range(0,218):
    print(i)
    #打开workbook
    data = xlrd.open_workbook('豆瓣'+str(i)+'.xls')
    # 获取sheet页
    table = data.sheet_by_name('sheet1')
    # 已有内容的行数和列数
    nrows = table.nrows

    for row in range(nrows):
        temp_list = table.row_values(row)
        if temp_list[0]!="用户" and temp_list[1]!="影评":
            data = []
            data.append([str(temp_list[0]),str(temp_list[1])])
            datalist.append(data)


import xlwt
# 创建一个workbook 设置编码
workbook = xlwt.Workbook(encoding = 'utf-8')
# 创建一个worksheet
worksheet = workbook.add_sheet('sheet1')

# 写入excel
# 参数对应 行, 列, 值
worksheet.write(0,0, label='用户')
worksheet.write(0,1, label='影评')

for i in range(0,len(datalist)):
    # 写入excel
    # 参数对应 行, 列, 值
    worksheet.write(i+1,0, label=str(datalist[i][0][0]))
    worksheet.write(i+1, 1, label=str(datalist[i][0][1]))

# 保存
workbook.save('豆瓣.xls')
