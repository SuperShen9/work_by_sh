# -*- coding: utf-8 -*-
# author:Super
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os,random
import pandas as pd
pd.set_option('expand_frame_repr',False)
os.chdir('C:\Users\Administrator\Desktop\Sample Data help')
df=pd.read_excel('text.xlsx')
# print df[df['country']=='English'].head(3)
df_china=df[df['country']=='CHINA']
df_english=df[df['country']=='English']
print list(df)
list_china=[]
for i in df_china['job title']:
    list_china.append(i)

list_english=[]
for i in df_english['job title']:
    list_english.append(i)

# print list_china,list_english

df_out=pd.DataFrame()
for j in range(20):
    df_out.at[j,'last']=random.choice(list_china)+random.choice(list_english)

print df_out




