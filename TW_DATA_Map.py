# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os,openpyxl
os.chdir('D:\zlianxi\TW_DATA_Map')
or_wb = openpyxl.load_workbook('data.xlsx')
sheet1 = or_wb.get_sheet_by_name('Sheet1')
sheet2 = or_wb.get_sheet_by_name('Sheet2')
wb = openpyxl.load_workbook('Industry_RE.xlsx')
ws = wb.get_sheet_by_name('data')
list1=[]
# list2=[]
for row in range(2, ws.max_row + 1):
    list1.append(ws['A' + str(row)].value)
    # list2.append(ws['B' + str(row)].value)
for row1 in range(2,sheet1.max_row+1):
    # if sheet1['C' + str(row1)].value !=None:
    sheet1['D' + str(row1)]=sheet1['C' + str(row1)].value.replace(' ','')
    for i in range(1,len(sheet1['D' + str(row1)].value)):
        if sheet1['D' + str(row1)].value[i:i+2] in list1:
            sheet1['E' + str(row1)] = sheet1['D' + str(row1)].value[:i].replace('國際','')
            sheet1['F' + str(row1)] = sheet1['D' + str(row1)].value[i:i+2].\
                replace('企業','').replace('股份','').replace('有限','').replace('公司','')
            break
if 'Find_M' in or_wb.get_sheet_names():
    or_wb.remove_sheet(or_wb.get_sheet_by_name('Find_M'))
    or_wb.create_sheet(index=2, title='Find_M')
else:
    or_wb.create_sheet(index=2,title='Find_M')
f_sheet=or_wb.get_sheet_by_name('Find_M')
f_sheet['A1']='标准公司名称'
f_sheet['B1']='标记'
f_sheet['C1']='匹配公司名称'
f_sheet['D1']='Flag'


key_dup={}
key_lob={}
for row2 in range(2,sheet2.max_row+1):
    if sheet2['A' + str(row2)].value != None:
        key_o = sheet2['A' + str(row2)].value
        val_o = '与leads重复'
        key_dup.setdefault(key_o, val_o)
    if sheet2['B' + str(row2)].value != None:
        key_l = sheet2['B' + str(row2)].value
        val_l = '与OPPT重复'
        key_dup.setdefault(key_l, val_l)
    if sheet2['C' + str(row2)].value != None:
        key_s = sheet2['C' + str(row2)].value
        val_s = '与SDR重复'
        key_dup.setdefault(key_s, val_s)
    if sheet2['D' + str(row2)].value != None :
        key_b = sheet2['D' + str(row2)].value
        val_b = 'LOB数据'
        key_lob.setdefault(key_b, val_b)
for row1 in range(2,sheet1.max_row+1):
    if sheet1['D' + str(row1)].value in key_dup.keys():
        sheet1['G' + str(row1)]=key_dup[sheet1['D' + str(row1)].value]
    else:
        if sheet1['F' + str(row1)].value != None:
            re_merge = sheet1['E' + str(row1)].value + sheet1['F' + str(row1)].value
            for ky in key_dup.keys():
                if re_merge in ky or ky in re_merge :
                    sheet1['G' + str(row1)] = key_dup[ky]
                    sheet1['H' + str(row1)] = ky
                    break
for row2 in range(2, sheet1.max_row + 1):
    if sheet1['D' + str(row2)].value in key_lob.keys():
        sheet1['I' + str(row2)] = sheet1['D' + str(row2)].value
    else:
        if sheet1['F' + str(row2)].value != None:
            re_merge2 = sheet1['E' + str(row2)].value + sheet1['F' + str(row2)].value
            for ky2 in key_lob.keys():
                if re_merge2 in ky2 or ky2 in re_merge2:
                    sheet1['J' + str(row2)] = ky2
                    break
i1 = 2
for row3 in range(2, sheet1.max_row + 1):
    if  sheet1['E' + str(row3)].value!= None \
        and sheet1['G' + str(row3)].value == None \
        and sheet1['H' + str(row3)].value == None :
        k_3=sheet1['E' + str(row3)].value
        if len(k_3)==2 or len(k_3)==3:
            for ky3 in key_dup.keys():
                if k_3 in ky3 :
                    f_sheet['A' + str(i1)] = sheet1['D' + str(row3)].value
                    f_sheet['C' + str(i1)] = ky3
                    f_sheet['D' + str(i1)] = key_dup[ky3]
                    sheet1['G' + str(row3)] = 'FIND'
                    i1+=1
                elif sheet1['G' + str(row3)].value == None:
                        sheet1['G' + str(row3)] = 'NON DUP'
for row3 in range(2, sheet1.max_row + 1):
    if sheet1['E' + str(row3)].value != None \
            and sheet1['I' + str(row3)].value == None\
            and sheet1['J' + str(row3)].value == None:
        k_4 = sheet1['E' + str(row3)].value
        if len(k_4) == 2 or len(k_4) == 3:
            for ky4 in key_lob.keys():
                if k_4 in ky4:
                    f_sheet['A' + str(i1)] = sheet1['D' + str(row3)].value
                    f_sheet['C' + str(i1)] = ky4
                    f_sheet['D' + str(i1)] = key_lob[ky4]
                    sheet1['I' + str(row3)] = 'FIND'
                    i1 += 1
                elif sheet1['I' + str(row3)].value == None:
                    sheet1['I' + str(row3)] = 'PL-NN'

sheet1.freeze_panes='A2'
sheet2.freeze_panes='A2'
f_sheet.freeze_panes='A2'
sheet1['D1']='标准公司名称'
sheet1['E1']='公司名简化'
sheet1['F1']='行业'
sheet1['G1']='Flag'
sheet1['H1']='Check'
sheet1['I1']='LOB'
sheet1['J1']='Check2'
or_wb.save('data.xlsx')