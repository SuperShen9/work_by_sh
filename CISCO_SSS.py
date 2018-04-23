# -*- coding: utf-8 -*-
# author:Super

import pandas as pd
import os
pd.set_option('expand_frame_repr', False)
# 切换到桌面需求文件目录
filepath = 'D:\super\CISCO SSS'
os.chdir(filepath)
for x, y, z in os.walk(filepath):
    # 读取文件
    df = pd.read_excel(z[0])

# 挑选出ECID不为空的、leads ID不重复的
df = df[df['Program Event Code'].notnull()]
df = df[df['Lead ID'] != '00Q340000220oWL']
df = df[df['Lead ID'] != '00Q34000021NrfJ']
df = df[df['Lead ID'] != '00Q34000021zgeF']
df = df[df['Lead ID'] != '00Q34000022TK55']
df = df[df['Lead ID'] != '00Q34000022TK4v']

# 导入 “活动名称” 和 “修改名称” 的2个表
df_event = pd.read_excel('sss_event.xlsx')
df_add = pd.read_excel('sss_zdd.xlsx')

# 合并导入的2个表
df = pd.merge(left=df, right=df_event,on='Program Event Code', how='left')
df = pd.merge(left=df, right=df_add,on='Lead ID', how='left')


# 同样的列合并2次，会复制原始数据2次
# print len(df[df['Program Event Code']==9661])
# os.chdir('C:\\Users\\Administrator\\Desktop')
# df.to_excel('CISCO Tracking.xlsx')
# exit()

# 去重没法使用（时间戳里的不同活动会有重复）
# df= df[df['Created Date']>pd.to_datetime('20171013')]
# df['dup']=df.duplicated(subset=['Company / Account','Last Name','Originating Marketing Pipeline'])
# df=df[df['dup']==False]

# 购买金额不为空，就除以1000；将修改名称列覆盖活动名称列
cdn_mql = (df['Originating Marketing Pipeline'].notnull())
df.loc[cdn_mql, 'Originating Marketing Pipeline2'] = df['Originating Marketing Pipeline']/1000
df.loc[df['add'].notnull(), 'Program View'] = df['add']

# 数据进行透视，并存入新的dataframe
df_s = pd.DataFrame()
df_s = df.groupby('Program View')
df_fin=pd.DataFrame()
df_fin['mql'] = df_s['Last Name'].count()
df_fin['mql_s'] = df_s['Originating Marketing Pipeline2'].sum()
df_fin['sql'] = df_s['Opportunity Name'].count()
df_fin['sql_s'] = df_s['Expected Total Value(000s) (converted)'].sum()

# 路径切换到桌面，导出数据
os.chdir('C:\\Users\\Administrator\\Desktop')
df_fin.to_excel('CISCO Tracking.xlsx')