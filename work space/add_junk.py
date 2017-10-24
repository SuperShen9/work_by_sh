# -*- coding: utf-8 -*-
# author:Super
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os,openpyxl,random
os.chdir('D:\zlianxi\\add_junk')
or_wb = openpyxl.load_workbook('AAAA.xlsx')
sheet1 = or_wb.active
hang = sheet1.max_row
list1=['~','/','《','!','@','#','$','%','^','&','*','(',')','>']
# ----------------------add part-------------------------------------
# list2=[]
# for j in range(50):
#     list2.append(random.choice(list1)+random.choice(list1))
#
#
# for i in range(2,hang+1):
#     # print sheet1['B'+str(i)].value.split()
#     sheet1['E' + str(i)] = random.choice(list2)+random.choice(list2).join(sheet1['B' + str(i)].value.split())+random.choice(list2)
#     sheet1['F' + str(i)] = random.choice(list2) + str(sheet1['C' + str(i)].value) + random.choice(list2)
#     sheet1['G' + str(i)] = random.choice(list2) + str(sheet1['D' + str(i)].value) + random.choice(list2)
# --------------------------------check-----------------------------------
for i in range(2,hang+1):
    for ii in list1:
        if sheet1['B' + str(i)].value!=None:
            if ii in sheet1['B' + str(i)].value:
                sheet1['C' + str(i)]='check'



or_wb.save('data.xlsx')