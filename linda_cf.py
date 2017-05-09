# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os,openpyxl
os.chdir('D:\zlianxi\linda_cf')
wbf = openpyxl.load_workbook('data.xlsx')
sheet=wbf.active
sheet2=wbf.get_sheet_by_name('Sheet2')
sheet3=wbf.get_sheet_by_name('Sheet3')
hang1 = sheet.max_row + 1
i=2
for row in range(2,hang1):
    sheet2['A' + str(i)] = sheet['A'+str(row)].value
    sheet2['B' + str(i)] = sheet['B'+str(row)].value
    sheet2['C'+str(i)]= sheet['C'+str(row)].value.replace('\n\n','|').replace('\n',' ')
    i+=1
j=2
for row in range(2,hang1):
    dodata = sheet2['C'+str(row)].value.strip().split('|')
    for dod in dodata:
        sheet3['A' + str(j)] = sheet2['A' + str(row)].value
        sheet3['B' + str(j)] = sheet2['B' + str(row)].value
        sheet3['C' + str(j)] = dod.encode('utf-8')
        j+=1

# dodata[0].encode('utf-8').strip()

wbf.save('data.xlsx')