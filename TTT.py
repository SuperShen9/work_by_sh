# -*- coding: utf-8 -*-
# author:Super
# import re
# num_Regex=re.compile(r'\d+')
# a={1:'aaaa',2:'rrr'}
# for i in range(1,3):
#     print 'b1b2'.replace(str(i),a[i])

# print  '路由器1'.replace('路由器', 'ROUTERS')
# a='as22dfgghhhgd'
#
# print a[2:4].isdigit()
# print a[2:3].isalpha()


# a = range(1,21)
# b = list(reversed(a))
# print b

# mo=num_Regex.findall('03 - 3283001 # 370')
# mo1=num_Regex.findall('033283001370')
# print mo
# print '-'.join(mo)
# print '-'.join(mo1)

# f = open("form.html",'w')
# message = """
# <html>
# <head>
# <meta charset="utf-8">
# <title>form</title>
# </head>
# <body>
#     <form>
#         <label>姓名</label>
#         <input type="text"/>
#         <br/>
#         <label>密码</label>
#         <input type="password"/>
#         <br/>
#         <label>性别</label>
#         <select>
#             <option>男</option>
#             <option>女</option>
#             <option>妖</option>
#         </select>
#         <br/>
#         <label>备注</label>
#         <br/>
#         <textarea rows="10" cols="30"></textarea>
#         <br/>
#         <button>提交</button>
#         <input type="submit" value="提交2"/>
#     </form>
# </body>
# </html>
# """
# f.write(message)
# f.close()

# f = open("div.html",'w')
# message = """
# <html>
# <head>
# <meta charset="utf-8">
# <title>div</title>
# </head>
# <body>
#     <div id="outer">
#         <div id="inner1">
#             <input type="text"/>
#     </div>
#     <div id="inner2">
#         <button>赞</button>
#     </div>
# </body>
# </html>
# """
# f.write(message)
# f.close()

# a='super beal made'
# b=a.split(' ')
# print len(b)
# a=input('please input:')
# if a==1:
#     print 'hehe'
# i=65
#
# for k in range(65,91):
#     print chr(i)
#     i+=1
#
#     spam = 65
detail_url=["https://www.qiushibaike.com/",]
for i in range(2, 14):
    url = "https://www.qiushibaike.com/8hr/page/" + str(i) + '/'
    detail_url.append(url)
print detail_url