# -*- coding: utf-8 -*-
from django.conf.urls import url
from .views import Users

urlpatterns=[
    url(r'', Users),
]

