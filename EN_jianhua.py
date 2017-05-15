# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os,openpyxl
os.chdir('D:\zlianxi\EN_jianhua')
or_wb = openpyxl.load_workbook('data.xlsx')
sheet1 = or_wb.get_sheet_by_name('Sheet1')
sheet2 = or_wb.get_sheet_by_name('Sheet2')
wb = openpyxl.load_workbook('SAS_RE.xlsx')
ws = wb.get_sheet_by_name('data')
list1=[]
list2=[]
for row in range(2, ws.max_row + 1):
    list1.append(ws['A' + str(row)].value)
    list2.append(ws['B' + str(row)].value)
for row1 in range(2,sheet1.max_row+1):
    sheet1['D' + str(row1)]=sheet1['C' + str(row1)].value.lower().strip()
    for i in range(1,len(list1)):
        if list2[i] ==None:
            kk=sheet1['D' + str(row1)].value.replace(str(list1[i]),'')
        else:
            kk = sheet1['D' + str(row1)].value.replace(str(list1[i]),str(list2[i]))
        sheet1['D' + str(row1)].value=kk
    sheet1['E' + str(row1)] = kk.strip()
    sheet1['D' + str(row1)] = sheet1['C' + str(row1)].value.lower().strip()
for row1 in range(2,sheet2.max_row+1):
    sheet2['D' + str(row1)]=sheet2['C' + str(row1)].value.lower().strip()
    for i in range(len(list1)):
        if list2[i] ==None:
            kk2=sheet2['D' + str(row1)].value.replace(str(list1[i]),'')
        else:
            kk2 = sheet2['D' + str(row1)].value.replace(str(list1[i]), list2[i])
        sheet2['D' + str(row1)].value=kk2.strip()
    sheet2['E' + str(row1)] = kk2.strip()
    sheet2['D' + str(row1)] = sheet2['C' + str(row1)].value.lower().strip()
# if 'Match_M' in or_wb.get_sheet_names():
#     or_wb.remove_sheet(or_wb.get_sheet_by_name('Match_M'))
#     or_wb.remove_sheet(or_wb.get_sheet_by_name('Find_M'))
#     or_wb.create_sheet(index=2, title='Match_M')
#     or_wb.create_sheet(index=3, title='Find_M')
# else:
#     or_wb.create_sheet(index=2,title='Match_M')
#     or_wb.create_sheet(index=3,title='Find_M')
# m_sheet=or_wb.get_sheet_by_name('Match_M')
# f_sheet=or_wb.get_sheet_by_name('Find_M')
# list_s1=[]
# list_s11=[]
# for row1 in range(2, sheet1.max_row + 1):
#     list_s1.append(sheet1['D' + str(row1)].value)
#     list_s11.append(sheet1['E' + str(row1)].value)
# list_s2=[]
# list_s22=[]
# for row2 in range(2, sheet2.max_row + 1):
#     list_s2.append(sheet2['D' + str(row2)].value)
#     list_s22.append(sheet2['E' + str(row2)].value)
# for m in list_s1:
#     if m in list_s2:
#         m_sheet['A'+str(i1)]=sheet1['A'+str(i1)]

sheet1.freeze_panes='A2'
sheet2.freeze_panes='A2'
sheet1['D1']='标准公司名称'
sheet1['E1']='AccountName_1'
sheet2['D1']='标准公司名称'
sheet2['E1']='AccountName_2'
or_wb.save('data.xlsx')