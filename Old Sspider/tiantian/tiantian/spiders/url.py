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
            detail_url = response.urljoin(i)
            req = Request(detail_url, self.parse_url)
            yield req

    def parse_url(self, response):
        i = response.xpath('//div[@class="fund-info"]/ul/li[3]/label/a/@href').extract()[0]
        d_url = response.urljoin(i)
        req = Request(d_url, self.parse_detail)
        yield req

    def parse_detail(self, response):
        for dd_url in response.xpath('//p[@class="table-content-title text-left"]/a/@href').extract():
            req = Request(dd_url, self.parse_dd)
            yield req

    def parse_dd(self, response):
        names = response.xpath('//div[@class="content_in "]/h4/span/text()').extract()[0]
        times = response.xpath('//div[@class="right jd "]/text()').extract()
        companys = response.xpath('//div[@class="right jd "]/a/text()').extract()[0]
        nums = response.xpath('//div[@class="gmleft gmlefts "]/span[2]/span/text()').extract()[0]
        pers = response.xpath('//div[@class="gmleft"]/span[2]/span[1]/text()').extract()[0]
        introduces = response.xpath('//div[@class="right ms"]/p/text()').extract()
        yield TiantianItem(name=names, time=times, company=companys, num=nums,
                           per=pers, introduce=introduces)
