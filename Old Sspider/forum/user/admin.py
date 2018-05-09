# -*- coding: utf-8 -*-
from django.contrib import admin
from .models import User_active

class Activate_Admin(admin.ModelAdmin):
    list_display = ('username','active_code','over_time')

admin.site.register(User_active,Activate_Admin)

