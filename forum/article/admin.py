# -*- coding: utf-8 -*-
from django.contrib import admin
from .models import Article

class ArticleAdmin(admin.ModelAdmin):
    list_display = ('block_id','title','content')

admin.site.register(Article)

