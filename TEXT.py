# -*- coding: utf-8 -*-
import os,openpyxl,datetime
from datetime import *
# os.chdir('D:\zlianxi\\text')
# wb = openpyxl.load_workbook('data.xlsx')
# sheet=wb.active
# time1=datetime.today()
# sheet['B2']=time1
# sheet['B3']=time1.year
# sheet['B4']=time1.month
# sheet['B5']=time1.day
# sheet['B6']=str(time1.year)+'-'+str(time1.month)+'-'+str(time1.day)
# wb.save('data.xlsx')

list1=['1','2','3','4']

for x in range(len(list1)):
    list1.remove(list1[x])
    print list1
    list1.append(list1[x])
    # d.append(x)


