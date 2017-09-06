# -*- coding: utf-8 -*-
# author:Super
import scrapy
class qiubaispider(scrapy.Spider):
    name = "qiubai"
    start_urls = [
        "https://www.qiushibaike.com/",
    ]
    def parse(self, response):
        print response.xpath('//div[@class="content"]').extract()
