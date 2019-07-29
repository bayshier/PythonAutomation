#! /usr/bin/env python
# -*- coding: utf-8 -*-
import os

from ArraysXml2ExcelI18n import arraysXmlToExcelI18n
from Excel2ArrayXmlI8n import excel2ArrayXml
from Excel2StringXmlI8n import excel2StringXml
#from Excel2StringXmlI8nAndTranslate import excel2StringXmlAndTrans
from StringXml2excelI18n import stringXmlToexcelI18n


def main():
    # arraysXmlToExcelI18n(os.getcwd())       # arrays.xml 转 Excel.xls
    # stringXmlToexcelI18n(os.getcwd())       # strings.xml 转 Excel.xls
    excel2ArrayXml()                          # Excel.xls 转 arrays.xml
    excel2StringXml()                         # Excel.xls 转 strings.xml


if __name__ == "__main__":
    main()
