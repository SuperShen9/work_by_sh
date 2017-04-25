# -*- coding: utf-8 -*-
import os,openpyxl
os.chdir('D:\zzsuper')
or_wb=openpyxl.load_workbook('data.xlsx')
wb=openpyxl.load_workbook('city.xlsx')
sheet=or_wb.create_sheet(index=0,title='data')
sheet1=or_wb.get_sheet_by_name('Sheet1')
sheetcity=wb.get_sheet_by_name('data')
hang=sheet1.get_highest_row()+1
citydata={}
for row in range(2,sheetcity.get_highest_row()+1):
    city=sheetcity['A'+str(row)].value
    post=sheetcity['B'+str(row)].value
    citydata.setdefault(city,post)
for i in range(2,hang):
    k=sheet1['B'+str(i)].value.strip().replace('  ',' ').replace('  ',' ').split(' ')
    if k[0].upper() in citydata.keys():
        print 1


