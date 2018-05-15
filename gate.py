# -*- coding: utf-8 -*-
# author:Super

from urllib2 import urlopen
import json, os
import pandas as pd
pd.set_option('expand_frame_repr', False)  # 当列太多时不换行

url = 'http://data.gateio.io/api2/1/pairs'

content = urlopen(url=url, timeout=15).read()
content = content.decode('utf-8')
con = content[2:-2].split('","')

# print(con)
# print(type(con))
# exit()

df=pd.DataFrame()

for x,y in enumerate(con):
    df.loc[str(x),'pairs']=y

os.chdir('C:\\Users\Administrator\Desktop')
df.to_csv('pairs.csv',index=False)
exit()

# 将数据转化为dataframe
json_data = json.loads(content)
df = pd.DataFrame(json_data, dtype='float')

# 对df进行处理
df = df[['ticker']].T

print(df)