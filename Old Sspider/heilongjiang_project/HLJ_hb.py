# -*- coding: utf-8 -*-
import os,openpyxl
os.chdir('D:\zlianxi\HLJ_hb')
wbf = openpyxl.load_workbook('data.xlsx')
sheet=wbf.active
hang1 = sheet.max_row + 1
spam = {}
for row in range(2, hang1):
    flag = sheet['D' + str(row)].value
    number = sheet['A' + str(row)].value
    spam.setdefault(flag, number)

wbf.create_sheet(index=2,title='dir_list')
d_sheet=wbf.get_sheet_by_name('dir_list')
i=2
for ky,ve in spam.items():
    d_sheet['A' + str(i)] = ky
    d_sheet['B' + str(i)] = ve
    i+=1

d_sheet.freeze_panes='A2'
d_sheet['A1']='小区名称'
d_sheet['B1']='ID'
wbf.save('data.xlsx')