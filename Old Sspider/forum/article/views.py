# -*- coding: utf-8 -*-
from django.shortcuts import render
from block.models import Block
from article.models import Article
from django.core.paginator import Paginator
def article_list(request,block_id):
    block_id = int(block_id)
    block=Block.objects.get(id=block_id)

    cut_page=5
    page_no = int(request.GET.get("page_no", "1"))
    all_article=Article.objects.filter(block=block,status=0).order_by("id")
    p=Paginator(all_article,cut_page)
    page=p.page(page_no)
    article_objs=page.object_list

    page_count=p.num_pages
    page_links=[i for i in range(1, page_count+1) if i>0 and i<=page_count]
    page_links1 = [i for i in range(page_no - 2, page_no + 3) if i > 0 and i <= page_count]
    page_small=page_links1[0]-1
    page_max=page_links1[-1]+1
    have_small=page_small>0
    have_max=page_max<page_count

    return render(request,'article_list.html',{'articles':article_objs,'b':block,
                                               'p':p,'page_no':page_no,'page_links':page_links,
                                               'page_small': page_small,'page_max':page_max,
                                               'have_small':have_small,'have_max':have_max,
                                               'page_count':page_count,'page_links1':page_links1})

    # page_no=int(request.GET.get("page_no","1"))
    # start_page=(page_no-1)*cut_page
    # end_page=page_no*cut_page
    # [start_page:end_page]