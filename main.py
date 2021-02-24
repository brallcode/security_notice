import os
import re
from openpyxl import Workbook
import requests
import json

import xlwt
from bs4 import BeautifulSoup
import urllib
import jsonpath
from openpyxl import Workbook
from openpyxl import load_workbook

url = "https://cert.360.cn/warning/searchbypage?length=6&start=0"
headers = {
    'Host': 'cert.360.cn',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:85.0) Gecko/20100101 Firefox/85.0',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate',
}

# html = requests.get(url=url, headers=headers)
add = urllib.request.Request(url=url, headers=headers)
html = urllib.request.urlopen(url=url, timeout=10)
# print(html.read())
html_json = json.loads(html.read())
# print(html_json)
id_list = jsonpath.jsonpath(html_json, "$..id")
# print(id_list)
# id_json = json.loads(html_json['data'])
# print(html_json['data'][0]['id'])

url_prefix = "https://cert.360.cn/warning/detail?id="
url_infos = []
for i in id_list:
    url_infos.append(url_prefix + i)
#    print(url_prefix + i)
# print(url_infos)

tmp_url = url_infos[3]
add2 = urllib.request.Request(url=tmp_url, headers=headers)
html2 = urllib.request.urlopen(url=tmp_url, timeout=10)
soup = BeautifulSoup(html2.read(), features="html.parser")
# 获取标题
title = soup.find("div", {"class": "detail-title-head"})
notice_name = title.get_text()
# print(title.get_text())

# 获取事件描述
notice_description = ''
# event_description = soup.find("div", {"class": "news-content"})
# print(event_description.get_text())
"""
# 遍历通报链接
for i in url_infos:
    add2 = urllib.request.Request(url=i, headers=headers)
    html2 = urllib.request.urlopen(url=i, timeout=10)
    print(html2.read())
"""

# 第一个h2 到 第二个 h2之间为事件描述
for i in soup.find("h2").next_siblings:
    if i.name == "h2":
        break
    # print(i.get_text())
    notice_description = notice_description + i.get_text()
# 获取重点漏洞 || 漏洞详情
info = ''
for i in soup.findAll("h2"):
    # 查找重点漏洞tag
    tag_zhongd = re.match(".*重点漏洞", i.string)
    # 获取终端漏洞描述
    if tag_zhongd is not None:
        for j in i.next_siblings:
            if j.name == "h2":
                break
            info = info + j.get_text()
        break
    # print(i.get_text())
    # 查找漏洞详情
    tag_loudongxq = re.match(".*漏洞详情", i.string)
    if tag_loudongxq is not None:
        for j in i.next_siblings:
            if j.name == "h2":
                break
            info = info + j.get_text()
# 获取修复建议
h2_jy = ''
for i in soup.findAll("h2"):
    h2_jy_name = re.match(".*修复建议", i.string)
    if h2_jy_name is None:
        continue
    for j in i.next_siblings:
        if j.name == 'h2':
            break
        h2_jy = h2_jy + j.get_text()
"""    
# 分别获取通用、临时修复建议
for i in soup.findAll("h2"):
    h2_jy = re.match(".*修复建议", i.string)
    if h2_jy is None:
        continue
    # print(i.next_sibling)
    # 通用修复建议
    h3_ty = i.next_sibling
    for i in h3_ty.next_siblings:
        if i.name == "h2":
            # 获取 临时建议 tag
            text_ty = re.match("临时修复建议", i.string)
            if text_ty is None:
                h3_ls = None
            else:
                h3_ls = i
            break
        # print(i.get_text())
    # 临时修复建议
    # if h3_ls is not None:
"""

# get release date: release_date
release_date_tag = soup.find("div", {"class": "detail-news-date"})
release_date_match = re.match("\d{4}-\d{1,2}-\d{1,2}", release_date_tag.string)
release_date = release_date_match.group()

# paqu xinxi list
notice = [release_date, notice_name, tmp_url, notice_description, info, h2_jy]
# print(notice)


# init excel
def excel_init(name):
    workbook = Workbook()
    row_1 = ['预警时间', '预警通报名称', '通报链接', '简述', '重点漏洞或漏洞详情', '修复建议']
    # worksheet = workbook.active
    ws1 = workbook.create_sheet('sheet1')
    ws1.append(row_1)
    workbook.save(name)


excel_name = '360cert预警通报统计.xlsx'
if not os.path.exists(excel_name):
    excel_init(excel_name)
workbook = load_workbook(excel_name)
worksheet = workbook['sheet1']
worksheet.append(notice)
workbook.save(excel_name)