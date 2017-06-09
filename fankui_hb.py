# -*- coding: utf-8 -*-
import os,openpyxl,xlrd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.styles.colors import RED
os.chdir('D:\superflag')
wbf = openpyxl.load_workbook('hehe.xlsx')
sheetcity = wbf.get_sheet_by_name('Sheet1')
hang1 = sheetcity.max_row + 1
spam = {}
for row in range(2, hang1):
    flag = sheetcity['A' + str(row)].value.lower()
    number = sheetcity['B' + str(row)].value
    spam.setdefault(flag, number)
os.chdir('D:\zlianxi\Fankui_hb')
k1=0
filepath=unicode('D:\zlianxi\Fankui_hb','utf-8')
for foldername,subfolder,excels in os.walk(filepath):
    baocun = openpyxl.Workbook()
    sheet = baocun.create_sheet(index=0, title='data')
    for excel in excels:
        wb = xlrd.open_workbook(excel)
        sheet1 = wb.sheets()[0]
        hang = sheet1.nrows
        lie = sheet1.ncols
        for k in range(lie):
            if sheet1.cell(0, k).value.lower() in spam.keys():
                kk = get_column_letter(spam[sheet1.cell(0, k).value.lower()])
                sheet[kk + '1'] = sheet1.cell(0, k).value
                sheet['A1'] = '来源'
                ft = Font(name='Arial', size=12, bold=True)
                ft1 = Font(name='Arial', size=12, bold=True, color=RED)
                sheet['A1'].font = ft1
                sheet[kk + '1'].font = ft
                j = 2
                for i in range(1, hang):
                    sheet['A' + str(j + k1)] = excel
                    sheet[kk + str(j + k1)] = sheet1.cell(i, k).value
                    j += 1
        k1 += hang - 1
sheet.freeze_panes='A2'
baocun.remove_sheet(baocun.get_sheet_by_name('Sheet'))
from datetime import *
time1=datetime.today()
hang2 = sheet.max_row + 1
for i in range(2,hang2):
    sheet['Y'+str(i)]=str(time1.year)+'-'+str(time1.month)+'-'+str(time1.day)
    sheet['AW' + str(i)] = str(time1.year) + '-' + str(time1.month) + '-' + str(time1.day)
    sheet['B' + str(i)] ='KDB00CI'
    if sheet['AI'+str(i)].value=='':
        sheet['AS' + str(i)]='N'
    else:
        sheet['AS' + str(i)] = 'Y'
    if sheet['AO'+str(i)].value=='':
        sheet['AU' + str(i)]='N'
    else:
        sheet['AU' + str(i)] = 'Y'
    if sheet['AV' + str(i)].value =='':
        sheet['AT' + str(i)] = 'N'
    else:
        sheet['AT' + str(i)] = 'Y'
    if sheet['R' + str(i)].value =='Hong Kong':
        sheet['R' + str(i)] ='香港'
        sheet['S' + str(i)]= '香港'

baocun.save('baocun.xlsx')










