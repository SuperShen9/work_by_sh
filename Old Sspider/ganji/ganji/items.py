# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy


class GanjiItem(scrapy.Item):
    # define the fields for your item here like:
    name = scrapy.Field()
    price = scrapy.Field()
    sale = scrapy.Field()
    rent =scrapy.Field()
    area = scrapy.Field()
    add = scrapy.Field()
    finishtime = scrapy.Field()
    liveroom = scrapy.Field()
    building = scrapy.Field()
    serve = scrapy.Field()
    pass
