# -*- coding: utf-8 -*-
# author:Super
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import pandas as pd
import time,os,shutil
import codecs
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\Administrator\Desktop')

df=pd.read_excel('sheet1.xlsx')

if os.path.exists('RUN'):
    shutil.rmtree('RUN')
os.makedirs('C:\\Users\Administrator\Desktop\\RUN')
os.chdir('C:\\Users\Administrator\Desktop\\RUN')

# df.shape[0]

for i in range(df.shape[0]):
    count=0
    for x in df.columns:
        count+=1
        val = df[x].loc[i]

        if isinstance(val, float):
            val = ''
        else:
            val = val
        fl = open('%s-%s-%s.txt' % (df['name'].loc[i],df['organization'].loc[i],df['webName'].loc[i]), 'a')

        if x == 'webName':
            fl.write('{\n')
            fl.write('\t\r"webUrl": "http://www.haodf.com/",\n')
            fl.write('\t\r"{}": "{}",'.format(x, str(val).decode('utf-8')))
            fl.write("\n")
        elif x == 'remark':
            fl.write('\t\r"remark":"",\n')
        elif x == 'isPerZone':
            fl.write('\t\r"{}": "{}",\n'.format(x, str(val).decode('utf-8')))
            fl.write('\t\r"MD": {\n')
            fl.write('\t\t\r"hospital": "",\n')
            fl.write('\t\t\r"hospitalLevel": "",\n')
            fl.write('\t\t\r"isTeachHos": "",\n')
            fl.write('\t\t\r"section": "",\n')
            fl.write('\t\t\r"docLevel": "",\n')
            fl.write('\t\t\r"goodAt": "",\n')
            fl.write('\t\t\r"hosAdminDuties": "",\n')
            fl.write('\t\t\r"hosAdminDutiesLevel": "",\n')
            fl.write('\t\t\r"hosTeachingPost": "",\n')
            fl.write('\t\t\r"hosTeachingPostLevel": "",\n')
            fl.write('\t\t\r"isAcademician": ""\n')
            fl.write('\t\r},\n')

            fl.write('\t\r"AT": [{\n')
            fl.write('\t\t\r"association": "",\n')
            fl.write('\t\t\r"associationLevel": "",\n')
            fl.write('\t\t\r"duty": "",\n')
            fl.write('\t\t\r"dutyLevel": "",\n')
            fl.write('\t\t\r"startAndEndTime": ""\n')
            fl.write('\t\r}],\n')

            fl.write('\t\r"JT": [{\n')
            fl.write('\t\t\r"journalName": "",\n')
            fl.write('\t\t\r"journalLevel": "",\n')
            fl.write('\t\t\r"duty": "",\n')
            fl.write('\t\t\r"dutyLevel": ""\n')
            fl.write('\t\r}],\n')

            fl.write('\t\r"AE": [{\n')
            fl.write('\t\t\r"time": "",\n')
            fl.write('\t\t\r"school": "",\n')
            fl.write('\t\t\r"major": "",\n')
            fl.write('\t\t\r"degree": ""\n')
            fl.write('\t\r}],\n')

            fl.write('\t\r"AW": [{\n')
            fl.write('\t\t\r"time": "",\n')
            fl.write('\t\t\r"name": "",\n')
            fl.write('\t\t\r"level": ""\n')
            fl.write('\t\r}],\n')

            fl.write('\t\r"WE": [{\n')
            fl.write('\t\t\r"startTime": "",\n')
            fl.write('\t\t\r"endTime": "",\n')
            fl.write('\t\t\r"workUnit": "",\n')
            fl.write('\t\t\r"title": ""\n')
            fl.write('\t\r}],\n')

            fl.write('\t\r"OC": [{\n')
            fl.write('\t\t\r"year": "",\n')
            fl.write('\t\t\r"territory": ""\n')
            fl.write('\t\r}]\n')

            fl.write("}")

        else:
            fl.write('\t\r"{}": "{}",\n'.format(x, str(val)))






