# -*- coding: utf-8 -*-
# Generated by Django 1.10 on 2017-11-27 14:36
from __future__ import unicode_literals

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('block', '0002_auto_20171122_2101'),
        ('article', '0001_initial'),
    ]

    operations = [
        migrations.CreateModel(
            name='Article',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('title', models.CharField(max_length=100, verbose_name='文章名称')),
                ('content', models.CharField(max_length=10000, verbose_name='文章信息')),
                ('status', models.IntegerField(choices=[(1, '常用'), (0, '待定'), (-1, '隐藏')], verbose_name='状态')),
                ('create_time_stamp', models.DateTimeField(auto_now_add=True, verbose_name='创建时间')),
                ('last_update_time_stamp', models.DateTimeField(auto_now=True, verbose_name='最新时间')),
                ('block_id', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='block.Block', verbose_name='板块id')),
            ],
            options={
                'verbose_name_plural': '文章',
                'verbose_name': '文章',
            },
        ),
        migrations.RemoveField(
            model_name='aeticle',
            name='block_id',
        ),
        migrations.DeleteModel(
            name='Aeticle',
        ),
    ]
