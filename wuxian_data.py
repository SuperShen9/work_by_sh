# -*- coding: utf-8 -*-
from urllib2 import urlopen
import json, os
import pandas as pd
import numpy as np
from datetime import timedelta
import datetime,time
from email.mime.text import MIMEText
from smtplib import SMTP
pd.set_option('expand_frame_repr', False)  # 当列太多时不换行


# # 定义抓取函数
# def get_candle_data(pair):
#     try:
# 抓取网址
symbol = 'btc_usd'
url = 'http://kx.18hezi.com/bitquote/v1/kline?exchange=bitfinex&symbol={}&size=100&type=1min'.format(symbol)

# 读取网址数据，并用json解析
content = urlopen(url=url, timeout=15).read()
content = content.decode('utf-8')
content = json.loads(content)


# 将数据转化为dataframe / 更改数据类型为float
df = pd.DataFrame(content, dtype=np.float)


# 重命名表的列名 / 转化时间戳  / 加8小时为北京时间
df['candle_begin_time'] = pd.to_datetime(df['timestamp'], unit='ms')
df['candle_begin_time_GMT8'] = df['candle_begin_time'] + timedelta(hours=8)

# 写入交易对
df['symbol'] = symbol.upper()
df = df[['symbol','candle_begin_time_GMT8','open','high','low','close','vol']]

print df
exit()

os.chdir('C:\Users\Administrator\Desktop')
df.to_excel('text.xlsx')


    #     return df
    # except KeyError:
    #     print('{}交易对失效,请及时修改'.format(symbol))






