# -*- coding: utf-8 -*-
# Generated by Django 1.10 on 2017-12-23 08:22
from __future__ import unicode_literals

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('user', '0001_initial'),
    ]

    operations = [
        migrations.RenameField(
            model_name='user_active',
            old_name='owner',
            new_name='username',
        ),
    ]
