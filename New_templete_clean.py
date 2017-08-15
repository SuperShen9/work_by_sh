# -*- coding: utf-8 -*-
# author:Super
import os,openpyxl
from TTT import *
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.styles.colors import RED
os.chdir('D:\zlianxi\New_templete_clean')
file = 'Clean_data.xlsx'
if os.path.exists(file):
    os.remove(file)
filepath = unicode('D:\zlianxi\New_templete_clean', 'utf-8')
list1=['Segment']
list2=['AM']

for foldername,subfolder,excels in os.walk(filepath):
    Clean_data = openpyxl.Workbook()
    sheet = Clean_data.create_sheet(index=0, title='data')
    wb=openpyxl.load_workbook(excels[0])
    sheet1 = wb.active
    hang = sheet1.max_row+1
    lie = sheet1.max_column+1
    print '数据量统计：%s'%(hang-2)
    i=0
    for k in range(1, lie):
        liebiao = get_column_letter(k)
        liebiao1=get_column_letter(k+i)
        i+=1
        for j in range(1, hang):
            sheet[liebiao1 + str(j)] = sheet1[liebiao + str(j)].value

    hang1 = sheet.max_row + 1
    lie1 = sheet.max_column + 1
    for kk in range(1 , lie1):
        lb=get_column_letter(kk)
        lb_1=get_column_letter(kk+1)
        for jj in range(2,hang1):
            if sheet[lb + '1'].value in list1:
                sheet[lb_1 + '1'] = '标准segment'
                sheet[lb_1 + str(jj)] = segment.get(sheet[lb + str(jj)].value)
            if sheet[lb + '1'].value in list2:
                sheet[lb_1 + '1'] = '标准AM'
                if sheet[lb + str(jj)].value!=None :
                    if sheet[lb + str(jj)].value in AM:
                        sheet[lb_1 + str(jj)] = sheet[lb + str(jj)].value
                    else:
                        sheet[lb_1 + str(jj)] ='XXXX'



        # sheet[liebiao1 + '1'] = sheet1[liebiao + '1'].value
        # for j in range(2, hang):
        #     if sheet1[liebiao + '1'].value=='COMPANY_NAME' or \
        #         sheet1[liebiao + '1'].value == 'LAST_NAME':
        #         sheet[liebiao + str(j)] = sheet1[liebiao + str(j)].value.strip()
        #
        #     else:
        #         sheet[liebiao + str(j)]=sheet1[liebiao + str(j)].value
        # if sheet1[liebiao + '1'].value.lower() in spam.keys():
        #     kk = get_column_letter(spam[sheet1[liebiao + '1'].value.lower()])
        #     sheet[kk + '1'] = sheet1[liebiao + '1'].value
        #     for j in range(1, hang):
        #         sheet['A' + str(j)] = excels[0]
        #         sheet[kk + str(j)] = sheet1[liebiao + str(j)].value
        #         j += 1
Clean_data.save('Clean_data.xlsx')