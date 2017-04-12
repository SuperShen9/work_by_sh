#coding:utf-8
with open('foo2.txt','r') as file:
    islogin = 1
    oklogin = 0
    while True:
        if oklogin == 0:
            if islogin == 4:
                print "the times is over,bye!"
                break
            else:
                account = raw_input('plz input your account:')
                passwd = raw_input('plz input your password:')
                for str1 in file:
                    user_account = str1.split(':')
                    print user_account
                    print user_account[1]
                    print str1
                    print
                    file_account = str1.split(':')[0]
                    file_passwd = str1.split(':')[1]
                    file_passwd = file_passwd[:-1]
                    if file_account == account and file_passwd == passwd:
                        oklogin = 1
                        break
                islogin += 1
                continue
        else:
            print 'login successful'
            break