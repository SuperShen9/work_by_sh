# -*- coding: utf-8 -*-
# author:Super

from re import findall
from urllib.request import urlopen
import os
url='https://mp.weixin.qq.com/s?__biz=MzI4MzM2MDgyMQ==&mid=2247485395&idx=1&sn=69fd710ccf17741c7fe0f7132a01aae5&chksm=eb8aac89dcfd259fe5742d69c5867da9019910741248f9c8fb1bfa4d5cbc9f0febfc77133b56&mpshare=1&scene=1&srcid=111725IOX2Pyvx1NG1GKgckF&pass_ticket=LgaTzmqQOHDYrmet8KWc6%2FFjdD05pzoI4Tx0iX%2Fo7WzYL3MDXsgic6ri83DAsU1M#rd'
with urlopen(url) as fp:
    content = fp.read().decode()
pattern = 'data-type="png" data-src="(.+?)"'
result = findall(pattern, content)
os.chdir('D:\pandas\weixin')
for index, item in enumerate(result):
    print(item)
    with urlopen(str(item)) as fp:
        with open(str(index)+'.png', 'wb') as fp1:
            fp1.write(fp.read())

