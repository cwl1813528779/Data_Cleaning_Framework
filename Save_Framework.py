# -*- coding: utf-8 -*-
"""
Created on Fri Sep 21 16:25:08 2018

@author: cwl

E-mail: 1813528779@qq.com

To: Sharing and gratitude are my attitude towards life
"""

import os
import time
import pandas as pd


class Save_Framework(object):
    """数据保存"""

    def __init__(self, Data_result):

        Dir = os.getcwd()
        self.Dir_result = Dir + os.sep + u'Result'
        self.Data_result = Data_result

    def Summary_Save_Framework(self):

        self.Critical_Save(self.Dir_result)
        self.Save_result(self.Data_result)

    def Save_result(self, Data_result):

        Dir_result = self.Dir_result + os.sep + \
            u'数据清洗-' + \
            time.strftime('%Y%m%d_%H%M.xlsx', time.localtime())

        Writer_result = pd.ExcelWriter(Dir_result)
        Data_result.to_excel(Writer_result,
                             sheet_name=u'数据清洗',
                             encoding='gb18030',
                             index=False)
        Writer_result.save()

    def Critical_Save(self, Dir):
        """保存文件路径判定"""
        if os.path.exists(Dir):
            pass
        else:
            os.mkdir(Dir)

# Data_result = pd.read_csv(u'e:\\data\\1.csv', keep_default_na=False, dtype=str, encoding='gb18030')
# Save_Framework(Data_result).Summary_Save_Framework()
