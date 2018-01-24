# -*- coding: utf-8 -*-
import pandas as pd
pd.set_option('expand_frame_repr',False)
# pd.set_option('max_colwidth',10)
df=pd.read_csv('sz000002.csv',
               nrows=10,
               skiprows=1,
               usecols=['交易日期','股票代码','收盘价','成交量','MACD_金叉死叉'],
               error_bad_lines=False,
               na_filter='NULL')

# print df['收盘价']==22.29
#
# print df[df['收盘价']>=25]
#
# print df[df.index <=4]
#
# print df.dropna(how='any')
#
# print df.fillna(value="无数据")
#
# df['MACD_金叉死叉'].fillna(value=df['交易日期'],inplace=True)
#
# print df
#
# print df.fillna(method='ffill')
#
# print df.fillna(method='bfill')
#
# print df.notnull()
#
# print df.isnull()
#
# # ascending=1,ignore_index=True
#
# df.reset_index(inplace=True)
#
# print df.sort_values(by=['交易日期'],ascending=1)
#
# print df.sort_values(by=['收盘价','成交量'],ascending=[0,1])
#
# df1=df.iloc[0:10][['股票代码','交易日期','收盘价']]
#
# df2=df.iloc[5:15][['股票代码','交易日期','成交量']]
#
# df3= df1.append(df2,ignore_index=True)
#
# df3.drop_duplicates(
#     subset=['股票代码','交易日期'],
#     keep='first',
#     inplace=True
# )
#
# print df3

# print df.rename(columns={'股票代码':'股票', 'MACD_金叉死叉':'金叉死叉'})

# print df.empty
# print pd.DataFrame().empty

# print df.T

# print df['股票代码'].str[:2]

# print df['股票代码'].str.upper()
# print df['股票代码'].str.lower()
# print df['股票代码'].str.len()
# print df['股票代码'].str.strip()
# print df['股票代码'].str.contain()
# print df['股票代码'].str.replace('sz','sh')

# print df['新浪概念'].str.split('；',expand=True)

df['交易日期']=pd.to_datetime(df['交易日期'])

# print df['交易日期']
# print df['交易日期'].dt.year
# print df['交易日期'].dt.dayofyear
# print df['交易日期'].dt.weekday
# print df['交易日期'].dt.weekday_name
# print df['交易日期'].dt.days_in_month
# print df['交易日期'].dt.is_month_end

# print df['交易日期']+pd.Timedelta(days=1)

# print df['交易日期']+pd.Timedelta(days=1)-df['交易日期']

# df['收盘价_3_mean']=df['收盘价'].rolling(3).mean()
#
# print df[['收盘价','收盘价_3_mean']]

df['收盘价_now']=df['收盘价'].expanding().mean()

# print df[['收盘价','收盘价_now']]

# df.to_csv('output.csv', encoding='gbk', index=False)
