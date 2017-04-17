import re
gzRegex=re.compile(r'3. .*')
text='3. what is the j num?'
t=gzRegex.search(text).group()
print t