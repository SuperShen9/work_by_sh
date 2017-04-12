# -*- coding: utf-8 -*-
str1 = ' abcdefgiiicde '
str2 = 'cde'
print str1.find(str2)
print str1.index(str2)
print str1.rfind(str2)#最后一次出现的位置
print str1.rindex(str2)#最后一次出现的位置
print str1.count(str2)#计数
print str1.replace(str2,'111')#替换的值必须为字符串
print str1.strip()#去两边空格 lstrip(left),rstrip(right)
print str1.split('c')
print 'H'.join(['ab','cd','ef'])
print str2.startswith('d')# 还有endswith()
print 'asd'.isalpha()
print '234'.isdigit()
print '   '.isspace()
print 'asd'.islower()
print 'ASD'.isupper()
print 'Asd'.istitle()
print '1.1'.isalnum()#判定是否为整数/digit是判定数字