# -*- coding: utf-8 -*-
from django.contrib import admin
from .models import Article

class ArticleAdmin(admin.ModelAdmin):
    list_display = ('block','title','owner','content',
                    'create_time_stamp','last_update_time_stamp')

admin.site.register(Article,ArticleAdmin)

