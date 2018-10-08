# -*- coding: utf-8 -*-
"""
Created on Fri Sep 14 14:07:46 2018

@author: cwl

E-mail: 1813528779@qq.com

To: Sharing and gratitude are my attitude towards life
"""

import pandas as pd


class Merge_Left(object):
    """左连接"""

    def __init__(self, Data_original, Data_left):

        self.Data_original = Data_original
        self.Data_left = Data_left

    def Summary_Deal(self, left_on=None, right_on=None,
                     on=None, how='left'):
        """
        汇总
        http://pandas.pydata.org/pandas-docs/stable/generated/pandas.merge.html
        参数
        ===============================================================
        :left_on: 左表主键
        :right_on: 右表主键
        :on: 两表相同字段
        :how: 数据连接方式,'left'未左连接,'right'未右连接;设置默认'left'

        return
        ==========================================
        :Data_result: DataFrame,执行连接后数据结果
        """
        self.Data_left = self.Data_left.drop_duplicates(right_on)
        Data_result = self.Merge_Just(
            self.Data_original,
            self.Data_left,
            how,
            left_on,
            right_on,
            on)
        # print(Data_result.columns)

        List_original = self.Data_original.columns.tolist()
        List_need = self.Data_left.columns.tolist()
        # print(List_original)
        # print(List_need)
        if right_on in List_need:
            List_need.remove(right_on)
        else:
            pass

        for column in List_need:
            if column in List_original:
                Data_result = self.Original_Column_Exist(Data_result, column)
            else:
                pass

        Data_result.fillna('', inplace=True)

        return(Data_result)

    def Merge_Just(self, Data_original, Data_left, how,
                   left_on, right_on, on):
        """
        数据连接
        参数
        ===================================
        :Data_original: 数据源左表
        :Data_left: 数据源右表
        :how: 'left'为左连接,'right'为右连接
        :left_on: 左表与右表连接字段
        :right_on: 右表与左表连接字段
        :on: 根据两表相同字段相连,可多个字段

        return
        =======================
        :Data_result: DataFrame
        """
        Data_result = pd.merge(
            Data_original,
            Data_left,
            how=how,
            left_on=left_on,
            right_on=right_on,
            on=on)

        return(Data_result)

    def Original_Column_Exist(self, Data_result, column):
        """
        两表相同字段处理
        参数
        ===========================
        :Data_result: 连接完的数据源
        :column: 两表相同的字段

        return
        =======================
        :Data_result: DataFrame
        """
        column_x = column + u'_x'
        column_y = column + u'_y'

        Index_need = Data_result[~Data_result[column_y].isnull()].index
        Data_result.loc[Index_need, [column_x]
                        ] = Data_result.loc[Index_need, column_y]

        del Data_result[column_y]

        Data_result = Data_result.rename(columns={column_x: column})

        return(Data_result)


# Data_original = pd.read_excel(u'E:\\python\\Merge\\测试.xlsx', dtype=str, keep_default_na=False)
# Data_Left = pd.read_excel(u'E:\\python\\Merge\\Data_Left.xlsx', dtype=str, keep_default_na=False)
# Data_result = Merge_Left(Data_original, Data_Left).Summary_Deal(left_on=u'婚博会id', right_on=u'婚博会id')
# Data_result.to_excel(u'测试结果.xlsx', encoding='gb18030', index=False)
