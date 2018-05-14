# -*- coding: utf-8 -*-
# author:Super
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import pandas as pd
import time,os,openpyxl
import codecs
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\Administrator\Desktop')
# #openpyxl模块
# wb = openpyxl.load_workbook('sheet1.xlsx')
# sheet = wb.get_sheet_by_name('Sheet1')
# print sheet['A2'].value
# print type(sheet['A2'].value.encode("gbk"))
# exit()

# # 查看text文件
# file=open('3.txt')
# lines=file.readlines()
# print lines
# f1=open('zhuan.txt', 'a')
# for i in lines:
#     f1.write(i)
# exit()

# file=open('2.txt')
# lines=file.readlines()
# # print lines
# for i in lines:
#     print i


# df=pd.read_excel('sheet1.xlsx')
#
# for i in range(1,df.shape[0]):
#     key = 'name'
#     val = df['NAME'].loc[i].encode('gbk')
#     # print val
#     fl = open('RUN.txt', 'a')
#     fl.write('\t\r"name\r"' + ':' +'\r"' + val+'\r"'+',')
#     fl.write("\n")


df = pd.read_excel('account without oppty_Becky_Yvonne.xlsx')
df_acc = pd.read_excel('HK_MODS list 20180510.xlsx')

df_1=pd.merge(left=df,
              right=df_acc,
              on=['Account ID'],
              how='left')
df_1.to_excel('out.xlsx')

