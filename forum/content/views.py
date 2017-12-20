# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.shortcuts import redirect
from block.models import Block
from article.models import Article
from article.forms import Articleform

def content(request,block_id):
    block_id = int(block_id)
    block = Block.objects.get(id=block_id)

    if request.method =='GET':
        return render(request,'content.html',{'b':block})

    else:
        form=Articleform(request.POST)
        if form.is_valid():
            article = form.save(commit=False)
            article.block=block
            article.status=0
            article.save()
            return redirect('/article/list/%s'%block_id)
        else:
            return render(request,'content.html',{'b':block,'form':form})

# 第一次的提交页面
# art_title =request.POST['art_title'].strip()
# art_content=request.POST['art_content'].strip()
#
# if not art_title or not art_content:
#     return render(request,'content.html',
#                   {'b': block,"error":"标题和内容不能为空",
#                    'art_title':art_title,'art_content':art_content})
#
# if len(art_title)>100 or len(art_content)>10000:
#     return render(request, 'content.html',
#                   {'b': block, "error": "标题和内容太长",
#                    'art_title': art_title, 'art_content': art_content})

# 第二次的提交页面
# if request.method =='GET':
#     return render(request,'content.html',{'b':block})
#
# else:
#     form=Articleform(request.POST)
#     if form.is_valid():
#         article = Article(block=block,title=form.cleaned_data["art_title"],
#                       status=0,content=form.cleaned_data["art_content"])
#         article.save()
#         return redirect('/article/list/%s'%block_id)
#     else:
#         return render(request,'content.html',{'b':block,'form':form})