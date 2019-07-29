#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xml.dom

import xlrd  # xlrd用来读excel文件
from googletrans import Translator



def excel2StringXmlAndTrans():
    # data = xlrd.open_workbook("ssss.xlsx")  # 打开excel文件
    data = xlrd.open_workbook("LauguageExcel.xls")  # 打开excel文件
    tab = data.sheet_by_index(0)  # 选择excel里面的Sheet
    nrows = tab.nrows  # 行数
    ncols = tab.ncols  # 列数

    for y in range(1, ncols):

        if type(tab.cell(0, y).value) == str and tab.cell(0, y).value != '':
            fileName = tab.cell(0, y).value
            languageType = y
            print(languageType, "out")

            dom1 = xml.dom.getDOMImplementation()  # 创建文档对象，文档对象用于创建各种节点。
            doc = dom1.createDocument(None, "resources", None)
            doc.documentElement.setAttribute('xmlns:xliff', 'urn:oasis:names:tc:xliff:document:1.2')
            top_element = doc.documentElement  # 得到根节点

            for x in range(0, nrows):
                # if x == 1000:
                # break
                sNode = doc.createElement('string')
                if type(tab.cell(x, 0).value) == float or x == 0:
                    # sNode.setAttribute('name', str(int(tab.cell(x, 0).value)))
                    continue  # 过滤Key为数字请求码
                else:
                    sNode.setAttribute('name', tab.cell(x, 0).value)
                    # 给这个节点加入文本，文本也是一种节点

                if type(tab.cell(x, languageType).value) == str:
                    print(languageType, "middle")



                    translator.translate('应用', dest='vi').text
                    text = doc.createTextNode(str(tab.cell(x, languageType).value))
                    # text = doc.createTextNode(Translator.translate(str(tab.cell(x, languageType).value), dest='fr').text)
                else:
                    text = doc.createTextNode(" ")
                # text = doc.createTextNode(content)
                sNode.appendChild(text)
                top_element.appendChild(sNode)

            print(languageType, "in")
            # print('strings'+ fileName +'.xml')
            f = open('strings' + fileName + '.xml', 'wb+')  # xml文件输出路径
            f.write(doc.toprettyxml().encode(encoding='utf-8'))
            f.close()
        else:
            break
