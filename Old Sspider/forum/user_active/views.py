# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.contrib.auth.models import User
from user.models import User_active
def activate(request,ac_code):
    ac_code=str(ac_code)
    if User_active.objects.get(active_code=ac_code):
        username=User_active.objects.get(active_code=ac_code).username
        user=User.objects.get(username=username)
        user.is_active=True
        user.save()
        return render(request,'active.html')
    else:
        return render(request, 'wrong.html')


