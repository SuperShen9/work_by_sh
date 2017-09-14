# -*- coding: utf-8 -*-
# author:Super
import scrapy
from qiubai.items import QiubaiItem
class qiubaispider(scrapy.Spider):
    name = "qiubai"
    start_urls = [
        "https://www.qiushibaike.com/",
    ]
    # def parse(self, response):
    #     print response.xpath('//div[@class="content"]').extract()
    def parse(self,response):
        for ele in response.xpath('//div[@class="article block untagged mb15"]'):
            authors = ele.xpath('./div[@class="author clearfix"]/a[2]/h2/text()').extract()
            contents = ele.xpath('.//div[@class="content"]/span/text()').extract()
            yield QiubaiItem(author=authors, content=contents)
