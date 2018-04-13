# -*- coding: utf-8 -*-
# author:Super
import pandas as pd
import time
from datetime import *
# df = pd.DataFrame({'col1': ['one', 'one', 'two', 'two', 'two', 'three', 'four'], 'col2': [1, 2, 1, 2, 1, 1, 1],
#                    'col3':['AA','BB','CC','DD','EE','FF','GG']},index=['a', 'a', 'b', 'c', 'b', 'a','c'])
# df['dup']=df.duplicated(subset=['col1','col2'])
# # print df
# print df[df['dup']==False]

# print pd.to_datetime('20171013')-datetime.today()

# def fx(a,b,c):
#     d = (a+b)*c
#     return a
#
# fx(1,2,3)

a=raw_input('请输入a,b,c:\n')
c=a.split(',')
d=(int(c[0])+int(c[1]))*int(c[2])

print '答案是：',d
