# coding=utf-8
with open('foo2.txt','a') as file:
    while 1:
        account = raw_input('plz input your account:')
        password = raw_input('plz input your password:')
        if account == '' or password == '':
            print 'plz input correct account or password'
            continue
        str1 = '%s:%s\n'%(account,password)
        print 'well done'
        break
    file.write(str1)
    file.flush()
