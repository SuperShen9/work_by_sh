# -*- coding: utf-8 -*-
from django.conf.urls import url
from .views import content

urlpatterns=[
    url(r'^content/(?P<block_id>\d+)',content),
]

