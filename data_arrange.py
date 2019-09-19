# -*- coding: utf-8 -*-
"""
Created on Sat May 25 10:02:52 2019

@author: Hu
"""
import xlrd
workbook=xlrd.open_workbook(整理.xlsx')
table = data.sheets()[0] #通过索引顺序获取
table.row_values(1)
table.col_values(1)