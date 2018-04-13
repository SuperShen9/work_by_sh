# -*- coding: utf-8 -*-
# author:Super

import pandas as pd
import os
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\\Administrator\\Desktop\\text')
df=pd.read_excel('FY17-FY18Q3 SFDC Leads Data_20180411 - HK & TW.xlsx')
df=df[df['Program Event Code'].notnull()]
df=df[df['Lead ID']!='00Q340000220oWL']
df=df[df['Lead ID']!='00Q34000021NrfJ']
df_event=pd.read_excel('sss_event.xlsx')
df_add=pd.read_excel('sss_zdd.xlsx')
df=pd.merge(left=df, right=df_event,on='Program Event Code', how='left')
df=pd.merge(left=df, right=df_add,on='Lead ID', how='left')
# 去重没法使用（时间戳里的不同活动会有重复）
# df= df[df['Created Date']>pd.to_datetime('20171013')]
# df['dup']=df.duplicated(subset=['Company / Account','Last Name','Originating Marketing Pipeline'])
# df=df[df['dup']==False]
cdn_mql=(df['Originating Marketing Pipeline'].notnull())
df.loc[cdn_mql,'Originating Marketing Pipeline2']=df['Originating Marketing Pipeline']/1000
df.loc[df['add'].notnull(),'Program View']=df['add']
df_s=pd.DataFrame()
df_s=df.groupby('Program View')
df_fin=pd.DataFrame()
df_fin['mql']=df_s['Last Name'].count()
df_fin['mql_s']=df_s['Originating Marketing Pipeline2'].sum()
df_fin['sql']=df_s['Opportunity Name'].count()
df_fin['sql_s']=df_s['Expected Total Value(000s) (converted)'].sum()
os.chdir('C:\\Users\\Administrator\\Desktop')
df_fin.to_excel('111.xlsx')