# -*- coding: utf-8 -*-
from django.db import models
from django.contrib.auth.models import User

class User_active(models.Model):
    username = models.ForeignKey(User, verbose_name="作者")
    active_code = models.CharField('激活码', max_length=100)
    over_time = models.DateTimeField('创建时间', auto_now=True)

    def __str__(self):
        return self.active_code
    class Meta:
        verbose_name='激活码'
        verbose_name_plural = '激活码'

