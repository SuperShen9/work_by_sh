# -*- coding: utf-8 -*-

from django.db import models
from block.models import Block

from django.contrib.auth.models import User

class Article(models.Model):
    block = models.ForeignKey(Block, verbose_name="板块索引")
    owner = models.ForeignKey(User,verbose_name="作者")
    title = models.CharField('文章名称', max_length=100)
    content = models.CharField('文章内容', max_length=10000)
    status = models.IntegerField(u'状态', choices=((1, '常用'), (0, '待定'), (-1, '隐藏')))

    create_time_stamp=models.DateTimeField('创建时间',auto_now_add=True)
    last_update_time_stamp=models.DateTimeField('最新时间',auto_now=True)

    def __str__(self):
        return self.title

    class Meta:
        verbose_name='文章'
        verbose_name_plural = '文章'

