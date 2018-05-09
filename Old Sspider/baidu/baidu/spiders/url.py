# -*- coding: utf-8 -*-
# author:Super
import scrapy
class Baidupider(scrapy.Spider):
    name = "baidu"
    start_urls = [
        "https://www.baidu.com/s?wd=%E9%91%AB%E9%91%AB%E5%88%B6%E9%A6%99",
    ]

    def parse(self, response):
        print response


