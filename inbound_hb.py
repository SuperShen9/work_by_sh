# -*- coding: utf-8 -*-
import os,openpyxl
from openpyxl.cell import get_column_letter
from openpyxl.styles import Font,Style
os.chdir('D:\superflag')
wbf = openpyxl.load_workbook('inbound_flag.xlsx')
sheetcity = wbf.get_sheet_by_name('Sheet1')
hang1 = sheetcity.max_row + 1
spam = {}
for row in range(2, hang1 + 1):
    flag = sheetcity['A' + str(row)].value
    number = sheetcity['B' + str(row)].value
    spam.setdefault(flag, number)
os.chdir('D:\\zlianxi\inbound_hb')
k1=0
for foldername,subfolder,excels in os.walk('D:\\zlianxi\inbound_hb'):
    baocun = openpyxl.Workbook()
    sheet = baocun.create_sheet(index=0, title='data')
    for excel in excels:
        wb=openpyxl.load_workbook(str(excel))
        sheet1 = wb.active
        hang = sheet1.max_row+1
        lie = sheet1.max_column+1
        for k in range(1,lie):
            liebiao=get_column_letter(k)
            if sheet1[liebiao+'6'].value in spam.keys():
                kk = get_column_letter(spam[sheet1[liebiao+'6'].value])
                sheet[kk + '1']=sheet1[liebiao + '6'].value
                fontobj = Font(name='Arial', size=12, bold=True)
                styleobj = Style(font=fontobj)
                sheet['A1'].style = styleobj
                sheet[kk + '1'].style = styleobj
                j = 2
                for i in range(7,hang):
                    sheet['A'+str(j+k1)] = str(excel)
                    sheet[kk+str(j+k1)] =sheet1[liebiao+str(i)].value
                    j+=1
        k1+=hang-2
sheet.freeze_panes='A2'
sheet['A1'] = '来源'
baocun.remove_sheet(baocun.get_sheet_by_name('Sheet'))
baocun.save('baocun.xlsx')
