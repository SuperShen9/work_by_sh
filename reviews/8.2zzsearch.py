import re
gzRegex=re.compile(r'\d+')
text='122313141d'
t=gzRegex.findall(text)
print t