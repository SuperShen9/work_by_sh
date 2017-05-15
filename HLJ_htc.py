# -*- coding: utf-8 -*-
import os,openpyxl
os.chdir('D:\zlianxi\HLJ_htc')
wbf = openpyxl.load_workbook('data.xlsx')
sheet=wbf.active
hang1 = sheet.max_row + 1
list1=['J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y',
        'Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL']
j = 0
for i in range(2,hang1):
    if sheet['C'+str(i)].value!= None:
        j=0
    if sheet['C'+str(i)].value== None:
        j += 1
        for y in list1:
            if sheet[y+str(i)].value!=None and sheet[y+str(i-j)].value==None:
                sheet[y + str(i - j)]=sheet[y + str(i)].value

wbf.save('data.xlsx')