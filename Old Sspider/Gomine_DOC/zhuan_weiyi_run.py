# -*- coding: utf-8 -*-
# author:Super
import pandas as pd
import time,os,shutil
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\Administrator\Desktop')
# #openpyxl模块
# wb = openpyxl.load_workbook('sheet1.xlsx')
# sheet = wb.get_sheet_by_name('Sheet1')
# print sheet['A2'].value
# print type(sheet['A2'].value.encode("gbk"))
# exit()

# # 查看text文件
# file=open('text.txt')
# lines=file.readlines()
# print lines
# for i in lines:
#     print i
# exit()

df=pd.read_excel('sheet1.xlsx')

if os.path.exists('RUN'):
    shutil.rmtree('RUN')
os.makedirs('C:\\Users\Administrator\Desktop\\RUN')
os.chdir('C:\\Users\Administrator\Desktop\\RUN')
df.shape[0]

for i in range(1):
    count=0
    for x in df.columns:
        count+=1
        val = df[x].loc[i]

        if isinstance(val, float):
            val = ''
        else:
            val = val.encode('utf-8')
        fl = open('%s-%s-%s.txt' % (df['name'].loc[i],df['organization'].loc[i],df['webName'].loc[i]), 'a')
        if count<=len(df.columns)-4:
            if x == 'webName':
                fl.write('{\n')
                fl.write('\t\r"webUrl": "http://www.chictr.org.cn/searchproj.aspx",\n')
                fl.write('\t\r"{}": "{}",'.format(x, str(val)))
                fl.write("\n")
            elif x == 'remark':
                fl.write('\t\r"remark":"",\n')
            else:
                fl.write('\t\r"{}": "{}",'.format(x, str(val)))
                fl.write("\n")
        elif count==len(df.columns)-3:
            fl.write('\t\r"info": {\n')
            fl.write('\t\t\r"name": "{}",\n'.format(str(val)))

        elif count==len(df.columns)-1:
            if val[-1]=='0':
                fl.write('\t\t\r"{}": "{}",\n'.format(x, str(val)[:4]))
            else:
                fl.write('\t\t\r"{}": "{}",\n'.format(x, str(val)[4:]))

        elif count == len(df.columns):
            fl.write('\t\t\r"{}": "{}"\n'.format(x, str(val)))
            fl.write('\t}\n')
            fl.write('}')
        else:
            fl.write('\t\t\r"{}": "{}",\n'.format(x, str(val)))





