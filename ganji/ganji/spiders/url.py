# -*- coding: utf-8 -*-
# author:Super
import scrapy
from ganji.items import GanjiItem
class ganjispider(scrapy.Spider):
    name = "ganji"
    start_urls = [
        "http://jinhua.ganji.com/xiaoqu/",
    ]

    def parse(self, response):
        # for xiaoqu in response.xpath('//div[@class="info-title"]/a/text()').extract():
        #     print xiaoqu
        xiaoqus = response.xpath('//div[@class="info-title"]/a/text()').extract()
        yield GanjiItem(xiaoqu=xiaoqus)

