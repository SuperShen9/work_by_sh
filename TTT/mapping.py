# -*- coding: utf-8 -*-
import os,openpyxl,pprint
os.chdir('D:\zsuper')
print 'Opening workbook……'
wb=openpyxl.load_workbook('city.xlsx')
sheetcity=wb.get_sheet_by_name('data')
citydata={}
print 'Leading citys……'

for row in range(2,sheetcity.get_highest_row()+1):
    city=sheetcity['A'+str(row)].value
    post=sheetcity['B'+str(row)].value
    citydata.setdefault(city,post)
# spam=sheetcity['A2'].value
# print sheetcity['A2'].value[1:2]
print sheetcity['A2'].value[-2:]
# print spam[:2]
# pprint.pprint(citydata.keys())
# if sheetcity['A2'] in sheetcity.column[1]:
#     print 1
# else:
#     print 2
