#!/usr/local/bin/python3
#构造表格，进行初始化

import xlrd
import xlwt
'''
import string
import numpy
import pandas
from xlutils.copy import copy

#读取大乐透开奖数据
workbook = xlrd.open_workbook(u'dlt-01.xls')
sheet_names = workbook.sheet_names()

sheet1 = workbook.sheet_by_name('results')

rows = sheet1.row_values(1)

#前区开奖结果数字
qian_array1 = rows[1]
qian_array2 = [0]*5

for i in range(0,5):
    qian_array2[0+i] = int(qian_array1[0+i*3:0+i*3+2])

#后区开奖结果数字
hou_array1 = rows[2]
hou_array2 = [0]*2

for j in range(0,2):
    hou_array2[0+j] = int(hou_array1[0+j*3:0+j*3+2])

#本期完整开奖结果
kaijian = qian_array2+hou_array2


print(rows)

print(qian_array1)
print(qian_array2)

print(hou_array1)
print(hou_array2)

print(kaijian)

#绘制excel表格，写入数据


#xlutils转换表格属性为可写入
#wbk = copy(workbook)
'''
#写入大乐透开奖数据
#前区
wbk = xlwt.Workbook()
sheet2 = wbk.add_sheet('qianqu',cell_overwrite_ok=True)

for i in range(0,35):
    #基本开奖结果统计
    sheet2.write(0,i+1,i+1)
    sheet2.write(i+1,0,i+1)
    
for j in range(1,36):
    for k in range(1,36):
        #表格各项赋初值
        sheet2.write(j,k,0)
    
    #补充次数统计
    #sheet2.write(37,i+1,sum(sheet2.cols[i+1]))
    #sheet2.write(i+1,37,sum(sheet2.rows[i+1]))


#后区
sheet3 = wbk.add_sheet('houqu',cell_overwrite_ok=True)
for i in range(0,12):
    sheet3.write(0,i+1,i+1)
    sheet3.write(i+1,0,i+1)
    
for j in range(1,13):
    for k in range(1,13):
        #表格各项附初值
        sheet3.write(j,k,0)

'''
#for i in range(1:47):

wbk.save(u'dlt-02.xls')

workbook2 = xlrd.open_workbook(u'dlt-02.xls')
wbk2 = copy(workbook2)
sheet4 = workbook2.get_sheet('qianqu')


temp = 0
#for q_i in range(1,2):
for l in range(0,5):
    for k in range(0,5):
        q_r = int(qian_array2[l])
        q_c = int(qian_array2[k])
        #int(temp)+1
        temp = 0 if l == k else sheet4.cell(q_r,q_c).value+1
        sheet2.write(q_r,q_c,temp)
'''
wbk.save('dlt-02.xls')





