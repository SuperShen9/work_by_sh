# -*- coding: utf-8 -*-
# author:Super
import pandas as pd
import matplotlib.pyplot as plt
pd.set_option('expand_frame_repr', False)
df=pd.read_hdf('D:\BaiduYunDownload\coin_quant_class_0518\coin_quant_class\data\class8\eth_1min_data.h5')

print df[df['candle_begin_time']>pd.to_datetime('20160606')].head(100)

# x=df['candle_begin_time']
# y=df['volume']
# plt.plot(x, y)
# plt.show()