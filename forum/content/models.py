# -*- coding: utf-8 -*-
from django.db import models
from block.models import Block
class Content(models.Model):
    block = models.ForeignKey(Block, verbose_name="板块索引")
    title = models.CharField('文章标题', max_length=100)
    content = models.CharField('文章内容', max_length=10000)
    status = models.IntegerField(u'状态', choices=((1, '常用'), (0, '待定'), (-1, '隐藏')))

    # def __str__(self):
    #     return self.title
    # class Meta:
    #     verbose_name = '内容'
    #     verbose_name_plural = '内容'

