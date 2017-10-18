# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy


class QiuzhiItem(scrapy.Item):
    # define the fields for your item here like:
    # name = scrapy.Field()
    job = scrapy.Field()
    company = scrapy.Field()
    exper = scrapy.Field()
    area = scrapy.Field()
    age = scrapy.Field()
    sex =scrapy.Field()
    salary = scrapy.Field()
    date = scrapy.Field()
    pass
