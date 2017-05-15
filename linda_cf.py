# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os,openpyxl
os.chdir('D:\zlianxi\linda_cf')
wbf = openpyxl.load_workbook('data.xlsx')
sheet=wbf.active

if 'sheet2' in wbf.get_sheet_names():
    wbf.remove_sheet(wbf.get_sheet_by_name('sheet2'))
    wbf.create_sheet(index=2, title='sheet2')
else:
    wbf.create_sheet(index=2, title='sheet2')

if 'sheet3' in wbf.get_sheet_names():
    wbf.remove_sheet(wbf.get_sheet_by_name('sheet3'))
    wbf.create_sheet(index=3, title='sheet3')
else:
    wbf.create_sheet(index=3, title='sheet3')

sheet2=wbf.get_sheet_by_name('sheet2')
sheet3=wbf.get_sheet_by_name('sheet3')

hang1 = sheet.max_row + 1
i=2
for row in range(2,hang1):
    sheet2['A' + str(i)] = sheet['A'+str(row)].value
    sheet2['B' + str(i)] = sheet['B'+str(row)].value
    sheet2['C'+str(i)]= sheet['C'+str(row)].value.replace('\n\n','|').replace('\n',' ')
    sheet2['D' + str(i)] = sheet['C' + str(row)].value
    i+=1
j=2
for row in range(2,hang1):
    dodata = sheet2['C'+str(row)].value.strip().split('|')
    for dod in dodata:
        sheet3['A' + str(j)] = sheet2['A' + str(row)].value
        sheet3['B' + str(j)] = sheet2['B' + str(row)].value
        sheet3['C' + str(j)] = dod.encode('utf-8').strip()
        j+=1

hang3 = sheet3.max_row + 1
list1=['E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X']
for row in range(2,hang3):
    if sheet3['C' + str(row)].value != '':
        dodata2 = sheet3['C' + str(row)].value.replace('联系人:','') \
            .replace('电话:', '').replace('传真:', '').replace('网址:', '') \
            .replace('地址:', '').replace('联系人备注:', '').replace(')', '').\
            replace('手机:', ' ').replace('(', '  ').replace('（', '  ') \
            .replace('）', '').replace('先生', ' 先生').replace('女士', ' 女士').split('  ')
        j1 = 0
        for dod2 in dodata2:
            sheet3[list1[j1]+str(row)]=dod2.encode('utf-8').strip()
            j1 += 1
    if sheet3['C' + str(row)].value == '':
        sheet3['A' + str(row)] = ''
        sheet3['B' + str(row)] = ''

jj=2
for row in range(2,hang1):
    dodata1 = sheet2['C' + str(row)].value.strip().split('|')
    for dod1 in dodata1:
        d=list(dodata1)
        d.remove(dod1)
        sheet3['D'+ str(jj)]='\n\n'.join(d)
        jj+=1

for row in range(2,hang3):
    if sheet3['C' + str(row)].value == '':
        sheet3['D' + str(row)] = ''

sheet3.freeze_panes = 'A2'
sheet3['A1']='ID'
sheet3['B1']='工程名称'
sheet3['C1']='业主 / 开发商'
sheet3['D1']='备注'
sheet3['E1']='分列start'

sheet2.freeze_panes = 'A2'
sheet2['A1']='ID'
sheet2['B1']='工程名称'
sheet2['C1']='去除换行符'
sheet2['D1']='原始数据列'

# dodata[0].encode('utf-8').strip()
wbf.save('data.xlsx')