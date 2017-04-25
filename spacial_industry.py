# -*- coding: utf-8 -*-
import os, openpyxl
os.chdir('D:\zlianxi\Spacial_industry')
or_wb = openpyxl.load_workbook('data.xlsx')
sheet1 = or_wb.get_sheet_by_name('Sheet1')
wb = openpyxl.load_workbook('city.xlsx')
sheetcity = wb.get_sheet_by_name('data')
sheethy = wb.get_sheet_by_name('data1')
hang = sheet1.max_row + 1
citydata = {}
hydata = {}
for row in range(2, sheetcity.max_row + 1):
    city = sheetcity['A' + str(row)].value
    post = sheetcity['B' + str(row)].value
    citydata.setdefault(city, post)
for row1 in range(2, sheethy.max_row + 1):
    hy = sheethy['A' + str(row1)].value
    hyid = sheethy['B' + str(row1)].value
    hydata.setdefault(hy, hyid)
for i in range(2,hang):
    if sheet1['B'+str(i)].value[:2] in citydata.keys() \
            and sheet1['B'+str(i)].value[-2:] in hydata.keys()\
            and sheet1['C'+str(i)].value[:2] in citydata.keys()\
            and sheet1['C'+str(i)].value[-2:]in hydata.keys():
        if hydata[sheet1['B'+str(i)].value[-2:]]<>hydata[sheet1['C'+str(i)].value[-2:]]:
            sheet1['D' + str(i)].value = '不匹配'
        elif sheet1['B'+str(i)].value[:2] not in sheet1['C'+str(i)].value \
            and sheet1['C'+str(i)].value[:2] not in sheet1['B'+str(i)].value:
            sheet1['D' + str(i)].value = '城市不同'
sheet1['D1'] = 'Flag'
sheet1.freeze_panes='A2'
or_wb.save('data.xlsx')



