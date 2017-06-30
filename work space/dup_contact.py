# -*- coding: utf-8 -*-
# author:Super
# python3版本
import os,openpyxl
os.chdir('D:\zlianxi\dup_contact')
wbf = openpyxl.load_workbook('data.xlsx')
sheet=wbf.active
hang = sheet.max_row + 1
lie = sheet.max_column-2
print (hang,lie)
list1=['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X']
for i in range(2,hang):
    d = ''
    for y in range(lie):
        a=sheet[list1[y]+'1'].value
        b=sheet[list1[y]+str(i)].value
        if b != None:
            # b = ''
            c = str(a) + ':' + str(b) + ';'
            d = d+' '+c
    sheet['B' + str(i)] = d
wbf.save('data.xlsx')