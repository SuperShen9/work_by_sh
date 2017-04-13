import re
def wocgex(woc):
    wocRegex=re.compile(r'[a-z]')
    woc1Regex=re.compile(r'[A-Z]')
    woc2Regex=re.compile(r'[0-9]')
    woc3Regex=re.compile(r'[^A-Za-z0-9]')
    k=wocRegex.search(woc)
    k1=woc1Regex.search(woc)
    k2=woc2Regex.search(woc)
    k3=woc3Regex.search(woc)
    if len(str(woc))<8:
        print ('The woc is to short')
    elif k<>None and k1<>None and k2<>None and k3==None:
        print ('The woc is Collect')
    else:
        print ('The woc is Un-standard')
woc=raw_input()
wocgex(woc)



