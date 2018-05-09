# -*- coding: utf-8 -*-
# author:Super
import scrapy
from scrapy.http import Request
from babybox.items import BabyboxItem

class babyboxspider(scrapy.Spider):
    name = "babybox"
    start_urls = [
        "http://www.babybox.com/",
    ]

    def parse(self, response):
        for i in response.xpath('//div[@class="top-navigation"]/a/@href').extract():
            detail_url = response.urljoin(i)
            req = Request(detail_url, self.parse_url)
            yield req

    def parse_url(self, response):
        for i in response.xpath('//center/a/@href').extract():
            d_url = response.urljoin(i)
            req = Request(d_url, self.parse_detail)
            yield req

    def parse_detail(self, response):
        for i in response.xpath('//center'):
            types = response.xpath('//div[@class="breadcrumbs"]//div[2]').xpath('string(.)').extract()
            names = i.xpath('./a/text()').extract()
            prices = i.xpath('./b/span/text()').extract()
            yield BabyboxItem(type=types, name=names, price=prices)

# 合并输出 产品加价格
# infos=response.xpath('//center/a/text()|//center/b/span/text()').extract()
# yield BabyboxItem(name=infos)

# 分开输出，结果会隔行，但是先保留吧
# for n in response.xpath('//center/a/text()').extract():
#     names = n
#     yield BabyboxItem(name=names)
# for p in response.xpath('//center/b/span/text()').extract():
#     prices = p
#     yield BabyboxItem(price=prices)
