# -*- coding: utf-8 -*-
import os,openpyxl
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
os.chdir('D:\zlianxi\Spacial_industry')
or_wb=openpyxl.load_workbook('data.xlsx')
wb=openpyxl.load_workbook('city.xlsx')
sheet=or_wb.create_sheet(index=0,title='data')
sheet1=or_wb.get_sheet_by_name('Sheet1')
sheetcity=wb.get_sheet_by_name('data')
sheethy=wb.get_sheet_by_name('data1')
hang=sheet1.max_row+1
citydata={}
hydata={}
for row in range(2,sheetcity.max_row+1):
    city=sheetcity['A'+str(row)].value
    post=sheetcity['B'+str(row)].value
    citydata.setdefault(city,post)
for row1 in range(2,sheethy.max_row+1):
    hy=sheethy['A'+str(row1)].value
    hyid=sheethy['B'+str(row1)].value
    hydata.setdefault(hy,hyid)
for i in range(2,hang):
    sheet['A' + str(i)].value = sheet1['A' + str(i)].value
    if sheet1['B'+str(i)].value[:3] in citydata.keys() :
        sheet['B'+str(i)].value=sheet1['B'+str(i)].value[3:]
    elif sheet1['B'+str(i)].value[:2] in citydata.keys():
        sheet['B' + str(i)].value = sheet1['B' + str(i)].value[2:]
    else:
        sheet['B' + str(i)].value = sheet1['B' + str(i)].value
for j in range(2,hang):
    if sheet1['C'+str(j)].value[:3] in citydata.keys():
        sheet['C'+str(j)].value=sheet1['C'+str(j)].value[3:]
    elif sheet1['C'+str(j)].value[:2] in citydata.keys():
        sheet['C' + str(j)].value = sheet1['C' + str(j)].value[2:]
    else:
        sheet['C' + str(j)].value = sheet1['C' + str(j)].value
for k in range(2,hang):
    if sheet1['B'+str(k)].value==sheet1['C'+str(k)].value:
        sheet['D' + str(k)].value = '完全匹配'
    elif sheet1['B'+str(k)].value in sheet1['C'+str(k)].value or sheet1['C'+str(k)].value in sheet1['B'+str(k)].value:
        sheet['D' + str(k)].value = '全称_Find'
    elif sheet['B'+str(k)].value in sheet['C'+str(k)].value or sheet['C'+str(k)].value in sheet['B'+str(k)].value:
        sheet['D' + str(k)].value = '简称_Find'
    elif sheet['B'+str(k)].value[-2:] in hydata.keys() or sheet['C'+str(k)].value[-2:] in hydata.keys():
        if sheet['B'+str(k)].value[-2:]==sheet['C'+str(k)].value[-2:]:
            if sheet['B' + str(k)].value[:2] == sheet['C' + str(k)].value[:2]:
                if sheet['B' + str(k)].value[-4:-2] == sheet['C' + str(k)].value[-4:-2]:
                    sheet['D' + str(k)].value = '基本匹配'
                else:
                    sheet['D' + str(k)].value = '行业错误'
            else:
                sheet['D' + str(k)].value = '地域错误'
        else:
            sheet['D' + str(k)].value = '不匹配'
        #     if sheet['B' + str(k)].value[-4:-2] == sheet['C' + str(k)].value[-4:-2]:
        #         if sheet['B'+str(k)].value[:2]==sheet['C'+str(k)].value[:2]:
        #             sheet['D' + str(k)].value = '基本匹配'
        #     else:
        #         sheet['D' + str(k)].value = '行业错误'
        # elif sheet['B' + str(k)].value[:2] == sheet['C' + str(k)].value[:2]:
        #     sheet['D' + str(k)].value = '地域甄别'

    elif sheet['B'+str(k)].value[:2]<>sheet['C'+str(k)].value[:2]:
        sheet['D'+str(k)].value='不匹配'
    else:
        sheet['D' + str(k)].value = '待定'
sheet['A1'] = 'ID'
sheet['B1'] = 'Company_Name'
sheet['C1'] = 'Dup_Name'
or_wb.save('data.xlsx')



