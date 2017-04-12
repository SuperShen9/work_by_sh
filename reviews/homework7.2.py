import re
def stripregex(x,y=None):
    if y==None:
        print (x.rstrip().lstrip())

    else:
        for i in y:
            reRegex=re.compile(i)
            mo=reRegex.sub('',x)
            x=mo
        print (x)
x=raw_input()
y=raw_input()
stripregex(x,y)


