# -*- coding: utf-8 -*-
# author:Super

import requests
from lxml import etree

comm_10 = []
for i in range(1,11):
    url='https://book.douban.com/subject/1084336/comments/hot?p='+str(i)
    r=requests.get(url).text
    s=etree.HTML(r)
    file = s.xpath('//div[@class="comment"]/p/text()')
    comm_10=comm_10+file


# #输出s对象
# print(s)

# #直接复制xpath路径
# print(s.xpath('/html/body/div[3]/div[1]/div/div[1]/div/div[2]/div/ul/li[1]/div[2]/p/text()')[0])

# #按标签优化提取
# print(s.xpath('//div[@class="comment"]/p/text()')[1])

# # 打印所有的评论
# list_comm=s.xpath('//div[@class="comment"]/p/text()')
# for comment in list_comm:
#     print(comment)

import os
os.chdir('C:\Users\Administrator\Desktop')

# with open('comm.txt', 'w') as f:
#    for i in file:
#        i=i.encode('utf-8')
#        f.write(i)

import pandas as pd
df = pd.DataFrame(comm_10)
df.to_excel('comm_10.xlsx',index=False)