# -*- coding: utf-8 -*-
from django.contrib import admin
from .models import Block


class BlockAdmin(admin.ModelAdmin):
    list_display = ('name','owner','desc')

admin.site.register(Block,BlockAdmin)

