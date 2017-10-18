# -*- coding: utf-8 -*-
# author:Super
import scrapy
from scrapy.http import Request
from qiuzhi.items import QiuzhiItem
class qiuzhispider(scrapy.Spider):
    name = "qiuzhi"
    start_urls = [
        "http://www.jh597.com/Person/Per_Search_Quick.shtml?x_suozaidi1=%BD%F0%BB%AA",
    ]
    def parse(self, response):
        for i in range(0, 100):
            if i == 0:
                detail_url = "http://www.jh597.com/Person/Per_Search_Quick.shtml?x_suozaidi1=%BD%F0%BB%AA"
            else:
                detail_url = "http://www.jh597.com/Person/Per_Search_Quick.shtml?PageNo=" + str(i)
            req = Request(detail_url, self.parse_url)
            yield req

    def parse_url(self, response):
        list1 = []
        for i in range(2, 80, 4):
            list1.append(i)
        count=0
        for row in response.xpath('//table[3]/./tr'):
            count+=1
            if count in list1:
                jobs = row.xpath('./td[2]/a/text()').extract()[0]
                companys=row.xpath('./td[3]/a/text()').extract()[0]
                expers=row.xpath('./td[4]/text()').extract()[0]
                ages = row.xpath('./td[5]/text()').extract()[0]
                sexs = row.xpath('./td[6]/text()').extract()[0]
                areas = row.xpath('./td[7]/text()').extract()[0]
                salarys = row.xpath('./td[8]/text()').extract()[0]
                dates = row.xpath('./td[9]/text()').extract()[0]
                yield QiuzhiItem(job=jobs, company=companys,exper=expers,area=areas,
                                 age=ages,sex=sexs,salary=salarys,date=dates)

        # print response.xpath('//table[3]/./tr[2]/td[2]/a/text()').extract()[0]
        # print response.xpath('//table[3]/./tr[2]/td[3]/a/text()').extract()[0]
        # print response.xpath('//table[3]/./tr[2]/td[4]/text()').extract()[0]
        # print response.xpath('//table[3]/./tr[2]/td[5]/text()').extract()[0]
        # print response.xpath('//table[3]/./tr[2]/td[6]/text()').extract()[0]
        # print response.xpath('//table[3]/./tr[2]/td[7]/text()').extract()[0]
        # print response.xpath('//table[3]/./tr[2]/td[8]/text()').extract()[0]
        # print response.xpath('//table[3]/./tr[2]/td[9]/text()').extract()[0]
        # response.xpath('//a[@class="line1"]/text()').extract()
        # response.xpath('//table[3]/./tr[2]/td/text()').extract()