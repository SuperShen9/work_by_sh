# -*- coding: utf-8 -*-
# author:Super
import scrapy
from scrapy.http import Request
from tiantian.items import TiantianItem
class tiantianspider(scrapy.Spider):
    name = "tiantian"
    start_urls = [
        "http://fund.eastmoney.com/company/default.html",
    ]
    def parse(self, response):
        for i in response.xpath('//div[@class="sencond-block"]/a/@href').extract():
            detail_urls=response.urljoin(i)
            yield TiantianItem(detail_url=detail_urls)
