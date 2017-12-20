# -*- coding: utf-8 -*-
# from django.http import HttpResponse
# 起初网页代码
# def index(request):
#     return HttpResponse("Hello 驻足五秒")

from django.shortcuts import render
from block.models import Block

def index(request):
    # block_infos = Block.objects.all().order_by('id')
    block_infos = Block.objects.filter(status__gte=0).order_by('id')
    return render(request,"index.html",{'blocks':block_infos})



# 板块简化输入blcok代码
# def index(request):
#     block_infos=[{'name':'python','desc':'learning','owner':'super'},
#                  {'name':'scrapy','desc':'learning','owner':'super'},
#                  {'name':'django','desc':'learning','owner':'super'}]
#     return render(request,"index.html",{'blocks':block_infos})

