# -*- coding: utf-8 -*-
"""
Created on Sat May 25 10:02:52 2019

@author: Hu
"""
import xlrd
import os

data=xlrd.open_workbook('arranged.xlsx')
table = data.sheets()[0] #通过索引顺序获取
# print(table.row_values(1))
col_info = table.col_values(0)
print(col_info)

##########################create fold#################
for col_name in col_info: #遍历所有文件
    try:
        # os.mkdir('./arranged_fold/%s'%col_name)
        os.mkdir('./arranged_fold/%s/ADC'%col_name)
        os.mkdir('./arranged_fold/%s/DWI' % col_name)
        os.mkdir('./arranged_fold/%s/Flair' % col_name)
        os.mkdir('./arranged_fold/%s/T1' % col_name)
        os.mkdir('./arranged_fold/%s/T1c' % col_name)
        os.mkdir('./arranged_fold/%s/T2' % col_name)
    except:
        pass
#######################################################



