# -*- coding: utf-8 -*-
from django.db import models

class Block(models.Model):
    name = models.CharField(u'板块名称',max_length=100)
    owner = models.CharField(u'管理员', max_length=100)
    desc = models.CharField(u'板块信息', max_length=100)
    status=models.IntegerField(u'状态',choices=((1,'常用'),(0,'待定'),(-1,'隐藏')))



    def __str__(self):
        return self.name

    class Meta:
        verbose_name='板块'
        verbose_name_plural='板块'

