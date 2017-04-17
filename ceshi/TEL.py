# -*- coding: utf-8 -*-
import os,openpyxl,re
os.chdir('D:\zlianxi\z1super')
or_wb=openpyxl.load_workbook('tel.xlsx')
sheet1=or_wb.get_sheet_by_name('Sheet1')
post=or_wb.get_sheet_by_name('post')
postdata={}
hang=sheet1.get_highest_row()+1
for row in range(2,post.get_highest_row()+1):
    ppp=post['A'+str(row)].value
    city=post['B'+str(row)].value
    postdata.setdefault(ppp,city)
list1=['886','86','852','853','886886']
for i in range(2,hang):
    sheet1['C1' ] = '标准电话'
    sheet1['D1'] = '标准城市'
    sheet1['E1'] = '处理过程'
    k=sheet1['B' + str(i)].value.find(list1[0])
    k1=sheet1['B' + str(i)].value.find(list1[1])
    k2 = sheet1['B' + str(i)].value.find(list1[2])
    k3 = sheet1['B' + str(i)].value.find(list1[3])
    k4 = sheet1['B' + str(i)].value.find(list1[4])
    if len(sheet1['B' + str(i)].value) < 8 and sheet1['B' + str(i)].value[0]=='+':
        sheet1['D' + str(i)] ='无效数据'
    elif -1<k4<2:
        sheet1['E' + str(i)]=sheet1['B'+str(i)].value[k+6:]
    elif -1<k<2:
        sheet1['E' + str(i)]=sheet1['B'+str(i)].value[k+3:]
    elif -1<k1<2:
        sheet1['E' + str(i)] = sheet1['B' + str(i)].value[k1 + 2:]
    elif -1 < k2 < 2:
        sheet1['C' + str(i)] = '852-'+sheet1['B' + str(i)].value[k2 + 3:]
        sheet1['D' + str(i)] ='香港'
    elif -1 < k3 < 2:
        sheet1['C' + str(i)] = '853-'+sheet1['B' + str(i)].value[k3 + 3:]
        sheet1['D' + str(i)] = '澳门'
    else:
        sheet1['E' + str(i)] = sheet1['B' + str(i)].value
for i in range(2,hang):
    if sheet1['E' + str(i)].value is not None:
        len1 = len(sheet1['E' + str(i)].value)
        if sheet1['E' + str(i)].value[:4] in postdata.keys() and len1>9:
            sheet1['C' + str(i)] = sheet1['E' + str(i)].value[:4] + '-' + sheet1['E' + str(i)].value[4:]
            sheet1['D' + str(i)] =postdata[str(sheet1['E' + str(i)].value[:4])]
        elif len1==11 and sheet1['E' + str(i)].value[0]=='1'\
                and sheet1['E' + str(i)].value[1]<>'1' and sheet1['E' + str(i)].value[1]<>'2':
            sheet1['C' + str(i)] = sheet1['E' + str(i)].value
            sheet1['D' + str(i)] ='大陆手机'
        elif len1==9 and sheet1['E' + str(i)].value[0]=='9':
            sheet1['C' + str(i)] = sheet1['E' + str(i)].value
            sheet1['D' + str(i)] ='台湾手机'
        elif sheet1['E' + str(i)].value[:3] in postdata.keys() and len1>9:
            if sheet1['E' + str(i)].value[0]=='0':
                sheet1['C' + str(i)]=sheet1['E' + str(i)].value[:3]+'-'+sheet1['E' + str(i)].value[3:]
                sheet1['D' + str(i)] = postdata[str(sheet1['E' + str(i)].value[:3])]
            else:
                sheet1['C' + str(i)] = '0'+sheet1['E' + str(i)].value[:3] + '-' + sheet1['E' + str(i)].value[3:]
                sheet1['D' + str(i)] = postdata[str(sheet1['E' + str(i)].value[:3])]
        elif sheet1['E' + str(i)].value[:2] in postdata.keys() and len1>9:
            if sheet1['E' + str(i)].value[0] == '0':
                sheet1['C' + str(i)] = sheet1['E' + str(i)].value[:2] + '-' + sheet1['E' + str(i)].value[2:]
                sheet1['D' + str(i)] = postdata[str(sheet1['E' + str(i)].value[:2])]
            else:
                sheet1['C' + str(i)] = '0' + sheet1['E' + str(i)].value[:2] + '-' + sheet1['E' + str(i)].value[2:]
                sheet1['D' + str(i)] = postdata[str(sheet1['E' + str(i)].value[:2])]

sheet1.freeze_panes='A2'
or_wb.save('tel.xlsx')