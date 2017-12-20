# -*- coding: utf-8 -*-
from django.conf.urls import url
from .views import Detail

urlpatterns=[
    url(r'^detail/(?P<article_id>\d+)', Detail),
]

