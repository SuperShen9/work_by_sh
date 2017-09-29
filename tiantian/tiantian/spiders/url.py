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
            detail_url=response.urljoin(i)
            req=Request(detail_url, self.parse_url)
            yield req

    def parse_url(self,response):
        names=response.xpath('//div[@class="ttjj-panel-title"]/p[1]/text()').extract()
        en_names=response.xpath('//div[@class="ttjj-panel-title"]/p[2]/text()').extract()
        yield TiantianItem(name=names,en_name=en_names)
