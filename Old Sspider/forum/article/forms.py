# -*- coding: utf-8 -*-

from django import forms
from article.models import Article
class Articleform(forms.ModelForm):
    class Meta:
        model=Article
        fields=['title','content']



    # art_title=forms.CharField(label="标题",max_length=100)
    # art_content=forms.CharField(label="内容",max_length=10000)

