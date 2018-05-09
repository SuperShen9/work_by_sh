# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.contrib.auth.models import User
from user.models import User_active
import uuid
from django.core.mail import send_mail
code = str(uuid.uuid4()).replace('-', '')

def Users(request):
    activate_link = 'http://%s/activate/%s' % (request.get_host(),code)

    if request.method =='GET':
        return render(request,'register.html')

    else:
        username=request.POST['username']
        email = request.POST['email']
        password = request.POST['password']
        password1 = request.POST['password1']

        activate_email = '''点击 <a href='%s'> 这里 </a>激活''' % activate_link
        send_mail(subject='[pythoner 网址] 激活邮件',
                  message='点击链接激活：%s' % activate_link,
                  html_message=activate_email,
                  from_email='2060633712@qq.com',
                  recipient_list=[email],
                  fail_silently=False)

        if password1 != password:
            return render(request, 'register.html', {'error': '密码不一致'})
        else:
            user = User.objects.create_user(username=username,
                                            email=email, password=password)
            user.is_active = False
            user.save()

            user_objs=User_active(username=user,active_code=code)
            user_objs.save()

            return render(request, 'user.html')





