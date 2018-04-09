# -*- coding: utf-8 -*-
# author:Super

import pandas as pd
import os
pd.set_option('expand_frame_repr',False)

os.chdir('C:\\Users\\Administrator\\Desktop\\text')
df=pd.read_excel('FY17-FY18Q3 SFDC Leads Data_20180404 - TW & HK.xlsx')

df=df[df['Program Event Code'].notnull()]

df_event=pd.read_excel('sss_event.xlsx')
# df_event.loc[df_event['Program Event Code']>90000,'add']='111'
# print df_event
# exit()
df_seg = pd.read_excel('sss_seg.xlsx')
df_st =  pd.read_excel('sss_st.xlsx')

df=pd.merge(left=df, right=df_event,on='Program Event Code', how='left')
df=pd.merge(left=df, right=df_seg,on='Lead Owner', how='left')
df=pd.merge(left=df, right=df_st,on='Lead Status', how='left')

cdn=(df['Lead Status2'].isnull())
cdn1=(df['Lead Status']=='2 Accepted-Mine/Channel')
cdn2=(df['Lost/Cancelled Reason']=='Cancelled - Opportunity created in error')
cdn3=(df['Opportunity Status']=='Active')
cdn31=(df['Opportunity Status']=='Booked')
cdn32=(df['Opportunity Status']=='Lost')
cdn33=(df['Opportunity Status']=='Cancelled')

cdn4=(df['Converted']==1)
cdn5=(df['Converted']==0)

df.loc[cdn & cdn2,'Lead Status2']='Cancelled Oppt'
df.loc[cdn & cdn3 ,'Lead Status2']='Converted SQL'
df.loc[cdn & cdn31,'Lead Status2']='Converted SQL'
df.loc[cdn & cdn32,'Lead Status2']='Converted SQL'
df.loc[cdn & cdn33,'Lead Status2']='Converted SQL'
df.loc[cdn1 & cdn4,'Lead Status2']='Accepted But Converted Issue'
df.loc[cdn1 & cdn5,'Lead Status2']='Accepted But Not Converted'

cdn_mql=(df['Originating Marketing Pipeline'].notnull())
df.loc[cdn_mql,'Originating Marketing Pipeline2']=df['Originating Marketing Pipeline']/1000

df.to_excel('123.xlsx')