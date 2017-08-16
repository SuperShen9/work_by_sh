# -*- coding: utf-8 -*-
# author:Super
import re
num_Regex=re.compile(r'\d+')
a={1:'aaaa',2:'rrr'}
for i in range(1,3):
    print 'b1b2'.replace(str(i),a[i])



# a = range(1,21)
# b = list(reversed(a))
# print b

# mo=num_Regex.findall('03 - 3283001 # 370')
# mo1=num_Regex.findall('033283001370')
# print mo
# print '-'.join(mo)
# print '-'.join(mo1)