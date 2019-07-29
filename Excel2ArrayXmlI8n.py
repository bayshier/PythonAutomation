#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xml.dom
import xlrd  # xlrd用来读excel文件


def excel2ArrayXml():
    # data = xlrd.open_workbook("ssss.xlsx")  # 打开excel文件
    data = xlrd.open_workbook("code.xls")  # 打开excel文件
    tab = data.sheet_by_index(0)  # 选择excel里面的Sheet
    nrows = tab.nrows  # 行数
    ncols = tab.ncols  # 列数

    for y in range(1, ncols):
#        type(tab.cell(0, y).value) == str and
        if tab.cell(0, y).value != '':
            fileName = tab.cell(0, y).value
            languageType = y
            print(languageType, "out")

            dom1 = xml.dom.getDOMImplementation()  # 创建文档对象，文档对象用于创建各种节点。
            doc = dom1.createDocument(None, "resources", None)
            top_element = doc.documentElement  # 得到根节点

            array = doc.createElement('string-array')
            array.setAttribute('name', 'premier_error_code')
            for x in range(0, nrows):
                # if x == 1000:
                # break
                sNode = doc.createElement('item')
#                print(str(tab.cell(x, 0).value))
                if str(tab.cell(x, 0).value) == "code":
                    continue
#                if type(tab.cell(x, 0).value) != float:
#                    continue
                # sNode.setAttribute('name', None)
                # if type(tab.cell(x, 1).value) == str:
                # sNode.setAttribute('name', tab.cell(x, 0).value)
                # else:
                # str = "" + tab.cell(x, 0).value
                # sNode.setAttribute('name', str)
                # 给这个节点加入文本，文本也是一种节点
#                if type(tab.cell(x, languageType).value) == str:
#                   text = doc.createTextNode(str(int(tab.cell(x, 0).value)) + '=' + tab.cell(x, languageType).value)
#                print(tab.cell(x, 0).value)
                text = doc.createTextNode(tab.cell(x, 0).value + '=' + tab.cell(x, languageType).value)
                    # sNode.appendChild(text)
#                else:
#                    text = doc.createTextNode(" ")
                sNode.appendChild(text)
                array.appendChild(sNode)

            top_element.appendChild(array)
            print(languageType, "in")
            # print('strings'+ fileName +'.xml')
            f = open('arrays' + fileName + '.xml', 'wb+')  # xml文件输出路径
            f.write(doc.toprettyxml(encoding='utf-8'))
#            f.write(doc.toprettyxml().encode(encoding='utf-8'))
            f.close()
        else:
            break
