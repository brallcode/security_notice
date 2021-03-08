"""
create update excel
2020-03-07
"""
import os

from openpyxl import Workbook, load_workbook

excel_name = '预警通报爬取统计.xlsx'
worksheet_name = '预警通报爬取统计'


def init_excel():
    row_1 = ['预警时间', '预警通报名称', '通报链接', '简述', '重点漏洞或漏洞详情', '修复建议']
    workbook = Workbook()
    global worksheet_name
    worksheet = workbook.create_sheet(worksheet_name)
    worksheet.append(row_1)
    global excel_name
    workbook.save(excel_name)


def excel_add(notices):
    global excel_name
    if not os.path.exists(excel_name):
        init_excel()
    workbook = load_workbook(excel_name)
    global worksheet_name
    worksheet = workbook[worksheet_name]
    for i in notices:
        worksheet.append(i)
    workbook.save(excel_name)
