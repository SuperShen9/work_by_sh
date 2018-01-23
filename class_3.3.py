# -*- coding: utf-8 -*-
# import pandas as pd
# pd.set_option('expand_frame_repr',False)
# # pd.set_option('max_colwidth',10)
# df=pd.read_csv('sz000002.csv',
#                # nrows=15,
#                skiprows=1,
#                usecols=['股票代码','交易日期','收盘价','成交量','MACD_金叉死叉'],
#                error_bad_lines=False,
#                na_filter='NULL')

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