# -*- coding: utf-8 -*-
from django.db import models
from article.models import Article
class Detail(models.Model):
    article = models.ForeignKey(Article, verbose_name="文章索引")
    title = models.CharField('文章标题', max_length=100)
    content = models.CharField('文章内容', max_length=10000)
    author = models.CharField('作者', max_length=10)

