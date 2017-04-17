
str1=['name','age','sex','game']
intro_new=open('intro_new.txt','w')
with open('introduce.txt','r') as file:
     for str2 in file:
         b=str2.split(' ')
         for i in b:
            if i in str1:
                replaceword=raw_input('what will replace %s:'%i)
                intro_new.write('%s '%replaceword)
            else:
                intro_new.write('%s '%i)
intro_new.close()
intro_new=open('intro_new.txt','r')
print intro_new.read()



