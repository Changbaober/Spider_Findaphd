# -*- coding: utf-8 -*-
"""
Created on Fri Dec  2 13:37:05 2022

@author: LENOVO
"""

# Automatically Get the Scholarship PhD projects
from urllib import request,parse
import ssl
from bs4 import BeautifulSoup
import time
import random
from ua_info import ua_list #使用自定义的ua池
import re
import pandas as pd
from openpyxl import load_workbook
import os
import xlwings as xw

class TiebaSpider(object):
    def __init__(self):
        self.url1='https://www.findaphd.com/phds/cross-subject/non-eu-students/?31M7qA4yx40'
        self.url='https://www.findaphd.com/phds/cross-subject/non-eu-students/?31M7qA4yx40&{}'

    def get_html(self,url):
        req=request.Request(url=url,headers={'User-Agent':random.choice(ua_list)})
        res=request.urlopen(req)
        html=res.read().decode("gbk","ignore")
        return html

    def parse_html(self,html):

        pattern1=re.compile(r'<h3 class.*?<a class=.*? href="(.*?)" title="PhD Research Project: (.*?) at (.*?)">*.*?',re.S)
        r_list1=pattern1.findall(html)
        pattern2=re.compile(r'<span class.*?/i>&nbsp;([0-9].*?)</span>',re.S)
        r_list2=pattern2.findall(html)
        data = []
        for i in range(0,len(r_list1)):
            d1 = r_list1[i]
            d = {"title": d1[1], "uni": d1[2], "ddl": r_list2[0], "link": 'https://www.findaphd.com'+d1[0]}
            data.append(d)
        return data
                
    def save_html(self,filename,page,data):
        insertData = pd.DataFrame(data)
        if not os.path.exists(filename):
            insertData.to_excel(filename, sheet_name='sheet1', index = False)
        else:
            df_old = pd.DataFrame(pd.read_excel(filename, sheet_name='sheet1'))
            row_old = df_old.shape[0] 
            book = load_workbook(filename)
            writer = pd.ExcelWriter(filename,engine='openpyxl') 
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            insertData.to_excel(writer, sheet_name='sheet1', startrow=row_old + 1, index=False, header=False)
            writer.save()
            writer.close()              

    def run(self):
        begin=int(input('start page：'))
        stop=int(input('end page：'))
        for page in range(begin,stop+1):
            if page == 0:
                url=self.url1
            else:
                PG=page + 1
                params={
                    'PG':PG,
                    }  
                params=parse.urlencode(params)
                url=self.url.format(params)
            html=self.get_html(url)
            data = self.parse_html(html)
            filename="D:\\PhD.xlsx" 
            self.save_html(filename,page,data)
            print('Successfully crawled the Page%d'%page)
            time.sleep(random.randint(1,10))

if __name__=='__main__':
    start=time.time()
    spider=TiebaSpider() 
    spider.run() 
    end=time.time()
    print('Totally run time:%.2f'%(end-start))  #爬虫执行时间    