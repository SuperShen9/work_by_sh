# -*- coding: utf-8 -*-
import os,openpyxl
os.chdir('D:\zlianxi\EN_fuzzy_yh')
or_wb=openpyxl.load_workbook('data.xlsx')
wb=openpyxl.load_workbook('zcity.xlsx')
# sheet=or_wb.create_sheet(index=0,title='data')
sheet1=or_wb.get_sheet_by_name('Sheet1')
sheetcity=wb.get_sheet_by_name('data')
hang=sheet1.max_row+1
citydata={}
for row in range(2,sheetcity.max_row+1):
    city=sheetcity['A'+str(row)].value
    post=sheetcity['B'+str(row)].value
    citydata.setdefault(city,post)

for i in range(2,hang):
    kb=sheet1['B'+str(i)].value.lower().replace('  ',' ').replace('  ',' ').split(' ')
    kd=sheet1['D'+str(i)].value.lower().replace('  ',' ').replace('  ',' ').split(' ')
    if kb[0] in citydata.keys():
        sheet1['F' + str(i)] = ' '.join(kb[1:])
    else:
        sheet1['F' + str(i)] = ' '.join(kb)

    if kd[0] in citydata.keys():
        sheet1['H' + str(i)] = ' '.join(kd[1:])
    else:
        sheet1['H' + str(i)] = ' '.join(kd)


sheet1.freeze_panes='A2'
or_wb.save('data.xlsx')
