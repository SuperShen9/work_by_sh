# -*- coding: utf-8 -*-
# author:Super

import pandas as pd
pd.set_option('expand_frame_repr', False)

path = 'C:\Users\Administrator\PycharmProjects\untitled\\babybox\\'

df = pd.read_csv(path+'baby.csv',
               error_bad_lines=False,
               na_filter='NULL')

df['real_price']=df['price'].shift(-1)

del df['price']

df = df[df['name'].notnull()]

df.drop_duplicates(
    subset=['name','real_price'],
    keep='first',
    inplace=True)

df.to_csv('finally_data.csv', index=False)