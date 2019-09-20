# -*- coding: utf-8 -*-
"""
Created on Sat May 25 10:02:52 2019

@author: Hu
"""
import xlrd
import xlwt
from xlutils.copy import copy
import os
import shutil
data=xlrd.open_workbook('arranged.xlsx')
wb = copy(data)  # 利用xlutils.copy下的copy函数复制
ws = wb.get_sheet(1)  # 获取表单0
table = data.sheets()[0] #通过索引顺序获取
# print(table.row_values(1))
col_info = table.col_values(0)
col_info.sort()
print(col_info[3:])
col_info = col_info[3:]
##########################create fold#################
for col_name in col_info: #遍历所有文件
    #######################THIS IS INITIAL########################
    # try:
    #     os.mkdir('./arranged_fold/%s'%col_name)
    #     # os.mkdir('./arranged_fold/%s/ADC'%col_name)
    #     # os.mkdir('./arranged_fold/%s/DWI' % col_name)
    #     # os.mkdir('./arranged_fold/%s/Flair' % col_name)
    #     # os.mkdir('./arranged_fold/%s/T1' % col_name)
    #     # os.mkdir('./arranged_fold/%s/T1c' % col_name)
    #     # os.mkdir('./arranged_fold/%s/T2' % col_name)
    # except:
    #     pass
    ###############################################################
    try:
        print(col_name)
        ro_pre = "F:/data/TCGA-GBM/%s/" % col_name
        print(ro_pre)
        houzhui = os.listdir(ro_pre)
        print(houzhui[0])
        ro_lat = "F:/data/TCGA-GBM/%s/%s" %(col_name,houzhui[0])
        print(ro_lat)
        ins_fol = os.listdir(ro_lat)
        print(ins_fol)

        for i in range(0,len(ins_fol)):
            ###########################################################arrange T1###################################################
            s1 = 'AX T1'
            result = s1 in ins_fol[i]
            if result==True:
                os.rename("F:/data/TCGA-GBM/%s/%s/%s" %(col_name,houzhui[0],ins_fol[i]), "F:/data/TCGA-GBM/%s/%s/T1" %(col_name,houzhui[0]))
                shutil.copytree("F:/data/TCGA-GBM/%s/%s/T1" %(col_name,houzhui[0]),
                            "F:/TCIA/arranged_fold/%s/T1"%col_name)
                ws.write(col_info.index(col_name), 2, 'F:/TCIA/arranged_fold/%s/T1'%col_name)  # 2 mean T1
                continue
            ###########################################################arrange FLAIR###################################################
            s2 = 'FLAIR'
            result = s2 in ins_fol[i]
            if result==True:
                os.rename("F:/data/TCGA-GBM/%s/%s/%s" %(col_name,houzhui[0],ins_fol[i]), "F:/data/TCGA-GBM/%s/%s/FLAIR" %(col_name,houzhui[0]))
                shutil.copytree("F:/data/TCGA-GBM/%s/%s/FLAIR" %(col_name,houzhui[0]),
                            "F:/TCIA/arranged_fold/%s/FLAIR"%col_name)
                ws.write(col_info.index(col_name), 5, 'F:/TCIA/arranged_fold/%s/FLAIR'%col_name)  # 5 mean FL
                continue
            ###########################################################arrange T2###################################################
            s3 = 'T2'
            result = s3 in ins_fol[i]
            if result == True:
                os.rename("F:/data/TCGA-GBM/%s/%s/%s" % (col_name, houzhui[0], ins_fol[i]),
                          "F:/data/TCGA-GBM/%s/%s/T2" % (col_name, houzhui[0]))
                shutil.copytree("F:/data/TCGA-GBM/%s/%s/T2" % (col_name, houzhui[0]),
                                "F:/TCIA/arranged_fold/%s/T2" % col_name)
                ws.write(col_info.index(col_name), 3, 'F:/TCIA/arranged_fold/%s/T2' % col_name)  # 3 mean T2
                continue
            ###########################################################arrange T1c###################################################
            s4 = 'T1 C'
            result = s4 in ins_fol[i]
            if result == True:
                os.rename("F:/data/TCGA-GBM/%s/%s/%s" % (col_name, houzhui[0], ins_fol[i]),
                          "F:/data/TCGA-GBM/%s/%s/T1c" % (col_name, houzhui[0]))
                shutil.copytree("F:/data/TCGA-GBM/%s/%s/T1c" % (col_name, houzhui[0]),
                                "F:/TCIA/arranged_fold/%s/T1c" % col_name)
                ws.write(col_info.index(col_name), 4, 'F:/TCIA/arranged_fold/%s/T1c' % col_name)
                continue
            ###########################################################arrange DWI###################################################
            s5 = 'DWI'
            s5_1 = 'Diffusion'
            result = s5 or s5_1 in ins_fol[i]
            if result == True:
                os.rename("F:/data/TCGA-GBM/%s/%s/%s" % (col_name, houzhui[0], ins_fol[i]),
                          "F:/data/TCGA-GBM/%s/%s/DWI" % (col_name, houzhui[0]))
                shutil.copytree("F:/data/TCGA-GBM/%s/%s/DWI" % (col_name, houzhui[0]),
                                "F:/TCIA/arranged_fold/%s/DWI" % col_name)
                ws.write(col_info.index(col_name), 6, 'F:/TCIA/arranged_fold/%s/DWI' % col_name)
                continue
            ###########################################################arrange adc###################################################
            s6 = 'ADC'
            result = s6 in ins_fol[i]
            if result == True:
                os.rename("F:/data/TCGA-GBM/%s/%s/%s" % (col_name, houzhui[0], ins_fol[i]),
                          "F:/data/TCGA-GBM/%s/%s/ADC" % (col_name, houzhui[0]))
                shutil.copytree("F:/data/TCGA-GBM/%s/%s/ADC" % (col_name, houzhui[0]),
                                "F:/TCIA/arranged_fold/%s/ADC" % col_name)
                ws.write(col_info.index(col_name), 7, 'F:/TCIA/arranged_fold/%s/ADC' % col_name)
                continue
    except:
        pass
wb.save('changed.xls')




