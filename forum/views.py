# -*- coding: utf-8 -*-
# author:Super
from django.shortcuts import render
def index(request):
    return render(request,'index.html')