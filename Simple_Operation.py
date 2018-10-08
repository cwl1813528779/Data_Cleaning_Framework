# -*- coding: utf-8 -*-
"""
Created on Mon Jul 30 17:13:13 2018

@author: cwl

E-mail: 1813528779@qq.com

To: Sharing and gratitude are my attitude towards life
"""

import six
#from numpy import str


class Index_Select(object):
    def __init__(self, Data_original, Data):
        """"拼接引用索引代码字符"""
        self.Data_original = Data_original
        self.Data = Data

    def Code_Join_Summary(self):
        """
        汇总
        return
        ===================================
        :Index_select: series,筛选出来的索引
        """
        Result_conditions = self.Code_Conditions_Join(self.Data)
        # six.print_(Result_conditions)
        Index_select = self.Data_original[eval(Result_conditions)].index

        return(Index_select)

    def Code_Conditions_Join(self, Data):
        """
        同个序号多个关联逻辑拼接
        参数
        ========================
        :Data: DataFrame

        return
        ===============
        Conditions: str
        """
        Conditions = ''
        for condition in Data.index:
            Data_related = Data.loc[condition]
            Conditions = Conditions + self.Code_Related_Join(Data_related)

        return(Conditions)

    def Code_Related_Join(self, Data_related,
                          Name_original=u'self.Data_original'):
        """
        单个关联逻辑拼接
        参数
        ===============================
        :Data_related: 同一序号逐行数据
        :Name_original: 数据源的变量名

        return
        ===============
        :Condition: str
        """
        Related_logic = {u'包含': [u'(', u'.str.contains("', '"))'],
                         u'不包含': ['~(', u'.str.contains("', '"))'],
                         u'等于': ['(', u'=="', '")'],
                         u'不等于': [u'(', u'!="', '")'],
                         u'大于': ['(', u'>"', '")'],
                         u'大于等于': [u'(', u'>="', '")'],
                         u'小于': [u'(', u'<"', '")'],
                         u'小于等于': [u'(', u'<="', '")'],
                         u'正则匹配': [u'(', u'.str.match("', '"))']
                         }

        if Data_related[u'字段关联逻辑'] in Related_logic.keys():
            Columns_condition = str(Data_related[u'序号关联逻辑']) + \
                Related_logic[Data_related[u'字段关联逻辑']][0] + \
                str(Name_original) + '["' + str(Data_related[u'关联字段']) + '"].astype(str)'

            Condition = Columns_condition + \
                Related_logic[Data_related[u'字段关联逻辑']][1] + \
                str(Data_related[u'关键字']) + \
                Related_logic[Data_related[u'字段关联逻辑']][2]
        elif Data_related[u'字段关联逻辑'] == u'字段相等':
            Columns_condition = str(Data_related[u'序号关联逻辑']) + u'(' + str(
                Name_original) + '["' + str(Data_related[u'关联字段']) + '"].astype(str)'
            Columns_aim = Name_original + \
                '["' + str(Data_related[u'关键字']) + '"].astype(str))'
            Condition = Columns_condition + u'==' + Columns_aim
        else:
            Columns_condition = str(Data_related[u'序号关联逻辑']) + u'(' + str(
                Name_original) + '["' + str(Data_related[u'关联字段']) + '"].astype(str)'

            Columns_aim = str(Name_original) + \
                '["' + str(Data_related[u'关联字段']) + '"].astype(str)'

            if Data_related[u'字段关联逻辑'] == u'All':
                Condition = Columns_condition + '==' + Columns_aim + ')'
            else:
                six.print_('Warning: The Related_logic: 【',
                           str(Data_related[u'字段关联逻辑']),
                           '】is not exist',
                           sep='')
                six.print_('The index will not be selected')
                Condition = Columns_condition + '!=' + Columns_aim + ')'

        return(Condition)

#Index_select = Index_Select(Data_original, Data_Operation).Code_Join_Summary()


class Simple_Operation(object):
    """简单操作"""

    def __init__(self, Data_original, Data):

        self.Data_original = Data_original
        self.Data_operation = Data.loc[Data.index[0]]  # 操作只取同一序号内第一行数据

    def Operation_Summary(self):
        """汇总"""
        Result_operation = self.Operation_Deal(self.Data_original,
                                               self.Data_operation)

        return(Result_operation)

    def Operation_Deal(self, Data_original, Data_operation):
        """
        操作处理
        参数
        =======================================================================
        :Data_original: DataFrame,数据源
        :Data_Operation: DataFrame,外接清洗指令
                         字段要求: 序号,序号关联逻辑,关联字段,字段关联逻辑,关键字,
                         操作字段,操作逻辑,关键字1,关键字2,拼接符号

        return
        =========================
        :Result_operation: series
        """
        # Column_operation = Data_operation[u'操作字段']
        # a = Data_original
        # b = Data_operation
        # c = Column_operation
        self.d = Data_operation[u'拼接符号']

        Dcit_operation = {u'前面增加': self.Add_front, u'后面增加': self.Add_last,
                          u'模糊替换': self.Replace_regex, u'精确替换': self.Replace,
                          u'批量模糊替换': self.Batch_replace_regex,
                          u'批量精确替换': self.Batch_replace,
                          u'完全覆盖': self.Cover, u'字段覆盖': self.Replace_column,
                          u'字段相加': self.Add_column, u'字段拼接': self.Cat_column,
                          u'字段数值相加': self.Add_number_column,
                          u'字段替换': self.Replace_column,
                          }

        if Data_operation[u'操作逻辑'] in Dcit_operation.keys():
            Func_operation = Dcit_operation[Data_operation[u'操作逻辑']]
            Result_operation = Func_operation(Data_original,
                                              Data_operation,
                                              Data_operation[u'操作字段'])
        else:
            Result_operation = Data_original[Data_operation[u'操作字段']]

        return(Result_operation)

    def Add_front(self, a, b, c):
        return(str(b[u'关键字1']) + a[c])

    def Add_last(self, a, b, c):
        return(a[c] + str(b[u'关键字1']))

    def Replace_regex(self, a, b, c):
        return(a[c].replace(str(b[u'关键字1']), str(b[u'关键字2']), regex=True))

    def Replace(self, a, b, c):
        return(a[c].replace(str(b[u'关键字1']), str(b[u'关键字2'])))

    def Batch_replace_regex(self, a, b, c):
        # six.print_(eval(b[u'关键字1']), eval(b[u'关键字2']))
        # six.print_(a[c].replace(eval(b[u'关键字1']), eval(b[u'关键字2']), regex=True))
        return(a[c].replace(eval(b[u'关键字1']), eval(b[u'关键字2']), regex=True))

    def Batch_replace(self, a, b, c):
        # six.print_(eval(b[u'关键字1']), eval(b[u'关键字2']))
        return(a[c].replace(eval(b[u'关键字1']), eval(b[u'关键字2'])))

    def Replace_column(self, a, b, c):
        return(a[b[u'字段辅助1']])

    def Cover(self, a, b, c):
        return(str(b[u'关键字1']))

    def Add_column(self, a, b, c):
        return(a[c].fillna('').astype(str) +
               a[b[u'字段辅助1']].fillna('').astype(str))

    def Cat_column(self, a, b, c):
        for item, value in enumerate(b[u'字段辅助1'].split('+')):
            try:
                if item == 0:
                    Result_operation = a[value].fillna('').astype(str)
                else:
                    Result_operation = Result_operation.str.cat(
                        a[value].fillna('').astype(str), sep=self.d)
            except BaseException:
                six.print_('The Column【', value, '】is not exist', sep='')

        return(Result_operation)

    def Add_number_column(self, a, b, c):
        for item, value in enumerate(b[u'字段辅助1'].split('+')):
            try:
                if item == 0:
                    Result_operation = a[value].fillna(
                        0).replace('', 0).astype(int)
                else:
                    Result_operation = Result_operation + \
                        a[value].fillna(0).replace('', 0).astype(int)
            except BaseException:
                six.print_('The Column:【', value, '】is not exist!', sep='')

        return(Result_operation)


def Summary_Operation(Data_original, Data_Operation):
    """
    数据清洗
    参数
    ========================================================================
    :Data_original: DataFrame,数据源
    :Data_Operation: DataFrame,外接清洗指令
                     字段要求: 序号,序号关联逻辑,关联字段,字段关联逻辑,关键字,
                     操作字段,操作逻辑,关键字1,关键字2,拼接符号

    return
    ==========================
    :Data_original: DataFrame
    """
    Levels_operation = Data_Operation[u'序号'].drop_duplicates()
    for operation in Levels_operation:
        try:
            Index_data = Data_Operation[Data_Operation[u'序号'] == str(
                operation)].index
            Data = Data_Operation.loc[Index_data]
            Column_operation = Data_Operation.loc[Index_data[0], u'操作字段']

            if not Data_original.columns.contains(Column_operation):
                Data_original.insert(
                    len(Data_original.columns), Column_operation, '')
            else:
                pass
            Index_select = Index_Select(
                Data_original, Data).Code_Join_Summary()
            Data_original.loc[Index_select, Column_operation] = Simple_Operation(
                Data_original.loc[Index_select], Data).Operation_Summary()
        except Exception as e:
            six.print_(u'Number:', operation, u'is error')
            six.print_(u'Error_Reason:', e)

    return(Data_original)

#import pandas as pd
#import six
#Data_original = pd.read_excel(u'e:\\python\\Simple_Operation\\测试.xlsx', keep_default_na=False, dtype=str)
#Data_Operation = pd.read_excel(u'E:\\python\\Simple_Operation\\Simple_Operation.xlsx', dtype=str, keep_default_na=False)
#Data_result = Summary_Operation(Data_original, Data_Operation)
#Data_result.to_excel(u'e:\\data\\1.xlsx', encoding='gb18030', index=False)
