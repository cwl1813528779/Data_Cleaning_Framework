# -*- coding: utf-8 -*-
"""
Created on Mon Aug 13 16:22:20 2018

@author: cwl

E-mail: 1813528779@qq.com

To: Sharing and gratitude are my attitude towards life
"""

import datetime
import Simple_Operation
from Merge_Left import Merge_Left
from File_Read_Traversal import File_Read_Traversal_Append as FR
from Save_Framework import Save_Framework


class Control_Data_Cleaning_Framework(object):
    """Project: Data_Cleaning_Framework"""

    def __init__(self):

        self.Data_original = FR(u'数据源').Summary_File_Read(
            changecsv=True, append_data=True)
        self.Data_select = FR(
            u'Configs\\1_数据筛选').Summary_File_Read(append_data=True)
        self.Data_merge_LoveId = FR(u'Configs\\2_Merge_LoveId').Summary_File_Read(
            sort_filename=True, TraversalExcel=True)
        self.Data_merge_phone = FR(u'Configs\\3_Merge_Phone').Summary_File_Read(
            sort_filename=True, TraversalExcel=True)
        self.Data_Operation = FR(u'Configs\\4_数据清洗').Summary_File_Read(
            sort_filename=True,
            TraversalExcel=True,
            add_file_name=True,
            add_read_time=True,
            append_data=True)

    def Control_Summary(self):
        """控制汇总"""
        print('Data is being filtered....')
        Data_result = self.Control_Select(self.Data_original, self.Data_select)
        print('Data is being matched....')
        Data_result = self.Control_Merge(Data_result,
                                         self.Data_merge_LoveId,
                                         self.Data_merge_phone)
        print('Data is being cleaned....')
        Data_result = self.Control_Cleaning(Data_result,
                                            self.Data_Operation)
        print('Data is being saved....')
        Save_Framework(Data_result).Summary_Save_Framework()

    def Control_Select(self, Data_original, Data_select):
        """数据筛选"""
        Levels_operation = Data_select[u'序号'].drop_duplicates()
        for item, operation in enumerate(Levels_operation):
            try:
                Index_data = Data_select[Data_select[u'序号'] == str(
                    operation)].index
                Data = Data_select.loc[Index_data]
                Index_select = Simple_Operation.Index_Select(
                    Data_original, Data).Code_Join_Summary()
                if item == 0:
                    Index_selects = Index_select
                else:
                    Index_selects = Index_selects.append(
                        Index_select).drop_duplicates()
            except Exception as e:
                print(u'Number:', str(operation), u' Is_Error', sep='')
                print(u'Error_Reason:', e)

        Data_result = Data_original.loc[Index_selects]

        return(Data_result)

    def Control_Merge(self, Data_result, Data_merge_LoveId,
                      Data_merge_phone):
        """数据左连接"""
        if u'婚博会id' in Data_result.columns:
            while True:
                try:
                    Data_LoveId = next(Data_merge_LoveId)
                    if len(Data_LoveId) != 0:
                        Data_LoveId.drop_duplicates([u'婚博会id'], inplace=True)
                        Data_result = Merge_Left(
                            Data_result,
                            Data_LoveId).Summary_Deal(
                            left_on=u'婚博会id',
                            right_on=u'婚博会id')
                    else:
                        pass
                except StopIteration:
                    break

        while True:
            try:
                Data_Phone = next(Data_merge_phone)
                if len(Data_Phone) != 0:
                    Data_Phone.drop_duplicates([u'Phone'], inplace=True)
                    for column_phone in [
                            u'爱人手机号', u'索票人手机号', u'新郎手机', u'新娘手机']:
                        if column_phone in Data_result.columns:
                            Data_result = Merge_Left(
                                Data_result,
                                Data_Phone).Summary_Deal(
                                column_phone,
                                u'Phone')
                            if u'Phone' in Data_result.columns:
                                del Data_result[u'Phone']
                            else:
                                pass
                else:
                    pass
            except StopIteration:
                break

        return(Data_result)

    def Control_Cleaning(self, Data_result, Data_Operation):
        """数据清洗"""
        for Columns_cat in [u'File_name_0', u'File_name_1', u'File_read_Time']:
            Data_Operation[u'序号'] = Data_Operation[u'序号'].str.cat(
                Data_Operation[Columns_cat], sep='_')

        Data_result = Simple_Operation.Summary_Operation(
            Data_result, Data_Operation)

        return(Data_result)


if __name__ == '__main__':
    print(__doc__)
    print(Control_Data_Cleaning_Framework.__doc__)
    print('========================================================')
    print('The program began to run....')
    print('----------------------------------')
    Start_Time = datetime.datetime.now()
    Control_Data_Cleaning_Framework().Control_Summary()
    End_Time = datetime.datetime.now()
    print('----------------------------------')
    print('Everything Done!')
    print('========================================================')
    print('Spend_Time', str(End_Time - Start_Time))
    input('Press any key to exit')
