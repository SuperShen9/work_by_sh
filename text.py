# -*- coding: utf-8 -*-
# author:Super
import pandas as pd
import time,os,openpyxl
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\Administrator\Desktop')
# #openpyxl模块
# wb = openpyxl.load_workbook('sheet1.xlsx')
# sheet = wb.get_sheet_by_name('Sheet1')
# print sheet['A2'].value
# print type(sheet['A2'].value.encode("gbk"))
# exit()

# # 查看text文件
file=open('1.txt')
lines=file.readlines()
print lines
# for i in lines:
#     print i
# exit()

# df=pd.read_excel('sheet1.xlsx')
#
#
# for i in range(1,df.shape[0]):
#     key = 'name'
#     val = df['NAME'].loc[i].encode('gbk')
#     # print val
#     fl = open('RUN.txt', 'a')
#     fl.write('\t\r"name\r"' + ':' +'\r"' + val+'\r"'+',')
#     fl.write("\n")



