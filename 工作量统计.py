# -*- coding: utf-8 -*-
# author:Super

import os, datetime
import pandas as pd
work_path = 'D:\gongzuoliang'

start_time = datetime.datetime.now()
type_list = []
type_dict = {}
df = pd.DataFrame()
for x, y, z in os.walk(work_path):
    for i in z:
        key = os.path.splitext(i)[1]
        if key=='':
            print os.path.splitext(i)[0]
            exit()
        type_dict.setdefault(key, 1)
        type_list.append(key)
k = 0
for dict in type_dict.keys():
    k += 1
    print('文件类型：{} - 数量：{} - 百分比：{}%'.format(dict,type_list.count(dict), round(float(type_list.count(dict))/float(len(type_list))*100),2))
    df.loc[k, 'Type'] = dict
    df.loc[k, 'count'] = type_list.count(dict)
    df.loc[k, 'percent'] = float(type_list.count(dict))/float(len(type_list))*100

os.chdir('C:\\Users\\Administrator\\Desktop')
df.to_csv('work.csv', index=False)
end_time = datetime.datetime.now()
print('\n程序运行时间:{}秒钟'.format((end_time - start_time).seconds))
