#!/usr/bin/env python
# -*- coding: utf-8 -*-
import datetime
import os
import xml.dom.minidom
import xlwt
import xlrd


# fileList传的第一个文件为默认的strings.xml文件，以该文件的KEY为基准去对应的匹配各个语言内容
def export2Excel(fileList):
    languageIndex = 0  # 语言标题的角标
    FILE_TITLE = 'Android-arrays翻译文档'
    # 获取当前时间
    today = datetime.date.today()
    # workbook = xlsxwriter.Workbook(FILE_TITLE + today.__str__() + '.xlsx')
    # worksheet = workbook.add_worksheet()

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet(FILE_TITLE + today.__str__(), cell_overwrite_ok=True)

    for file in fileList:
        languageIndex += 1

        # 使用minidom解析器打开 XML 文档
        DOMTree = xml.dom.minidom.parse(file)
        collection = DOMTree.documentElement
        # resources = collection.getElementsByTagName("resources")[0]
        strings = collection.getElementsByTagName("string-array")
        # strings = string_array.getElementsByTagName("string")

        # 从第一个单元格开始，行和列的索引均为0
        row = 1
        col = 0
        # 写标题
        worksheet.write(0, 0, "键")
        worksheet.write(0, languageIndex, os.path.splitext(file)[0])
        # 迭代数据并逐行写入
        # fo    r cost in (expenses):
        for string in strings:
            for item in string.getElementsByTagName("item"):

                if len(item.childNodes[0]) > 0:
                    # 默认的语言
                    data = item.childNodes[0].data.split('=')
                    if languageIndex == 1:
                        print(languageIndex, item.childNodes[0].data)
                        # worksheet.write(row, col, q.getAttribute("name"))  # 写KEY
                        worksheet.write(row, col, data[0])  # 写KEY
                        worksheet.write(row, languageIndex, data[1])  # 写VALUE
                        row += 1
                    # 其它语言以默认语言的KEY为标准填充各种的VALUE
                    else:
                        dataInner = xlrd.open_workbook(FILE_TITLE + today.__str__() + '.xls')  # 打开excel文件
                        tab = dataInner.sheet_by_index(0)  # 选择excel里面的Sheet
                        narrows: int = tab.nrows  # 行数
                        for innerRow in range(1, narrows):
                            print(str(tab.cell(innerRow, 0).value))
                            if str(int(tab.cell(innerRow, 0).value)) == data[0]:
                                # worksheet.write(innerRow, col, q.getAttribute("name"))  # 写KEY
                                worksheet.write(innerRow, languageIndex, data[1])  # 写VALUE
                            else:
                                continue  # 其它语言的KEY如果多的不添加

        workbook.save(FILE_TITLE + today.__str__() + '.xls')  # 保存文件
    print(FILE_TITLE + today.__str__() + '.xlsx', '导出成功')


# 获取该文件夹下的各个strings.xml文件
def file_name(file_dir):
    language_index = 0
    fileList = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.xml' and os.path.splitext(file)[0].find('arrays') != -1:
                print(file)
                language_index += 1
                fileList.append(file)
    export2Excel(fileList)


file_name(os.getcwd())
