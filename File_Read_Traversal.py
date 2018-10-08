# -*- coding: utf-8 -*-
"""
Created on Thu Aug 16 13:49:44 2018

@author: cwl

E-mail: 1813528779@qq.com

To: Sharing and gratitude are my attitude towards life
"""

from io import open
import os
import time
import re
import win32com.client
import six
import numpy as np
import pandas as pd


def File_Read_Decorator(File_Read_fun):
    def Decorator(*args, **kwargs):
        try:
            return(File_Read_fun(*args, **kwargs))
        except Exception as e:
            six.print_(e)
            # exit()

    return(Decorator)


class Fotmat_Change(object):
    """xlsx|xls|xlsb|xlsm is converted to csv"""

    def __init__(self, Dir_file):
        self.Dir_file = Dir_file
        Time_now = np.str(time.time())
        self.Dir_change = os.getcwd() + os.sep + Time_now
        self.Dir_save = self.Dir_change + os.sep + Time_now + u'.csv'

    def Summary_Deal(self):
        """
        汇总

        :return
        ==================================
        :self.Dir_save: 文件路径所在文件夹
        :self.Dir_change: 文件路径
        """
        self.Critical_Deal(self.Dir_change)
        six.print_(
            re.findall('\.(\d*\w*\d*)$', self.Dir_file)[0],
            u'is being converted to csv....')
        self.Read_Deal(self.Dir_file, self.Dir_change)

        return(self.Dir_save, self.Dir_change)

    def Read_Deal(self, Dir_file, Dir_change):
        """
        xlsx/xls/xlsb/xlsm调用系统Micro Excel转csv
        参数
        =========================
        :Dir_file: 文件读取路径
        :Dir_change: 文件存储路径

        return
        ======
        :None
        """
        Excel = win32com.client.Dispatch('Excel.Application')
        Excel.Visible = 0
        Excel.DisplayAlerts = 0  # 不在前台显示文档及错误
        Data_original = Excel.Workbooks.Open(Dir_file)
        Data_original.SaveAs(self.Dir_save, FileFormat=6)
        Data_original.Close()
        try:
            Excel.ActiveWorkbook.ActiveSheet.Cells(1, 1).Value
            Excel.Visible = 1
        except BaseException:
            Excel.Application.Quit()

    def Critical_Deal(self, Dir_change):
        """
        路径判断,不存在则创建文件夹
        参数
        ==========================
        :Dir_change: 文件读取路径

        return
        ======
        :None
        """
        if os.path.exists(Dir_change):
            pass
        else:
            os.mkdir(Dir_change)


class File_Read_Traversal_Append(object):
    """文件夹文件读取"""

    def __init__(self, File_name):

        self.File_name = File_name

    def Summary_File_Read(self, unfill=False, TraversalExcel=False,
                          changecsv=False, mydtype=six.text_type, append_data=False,
                          sort_filename=False, add_file_name=False,
                          add_read_time=False):
        """
        汇总
        参数
        ============================================================================
        :File_name: 当前工作环境下的文件路径
        :unfill: 布尔值,False将空值替换'';True则不替换;默认False
        :TraversalExcel: 布尔值,True则遍历工作簿下的每个非空工作表;默认False
        :changecsv: 布尔值,调用微软Excel转csv,非常规方法,尽量不使用;默认False
        :mydtype: 字段格式,默认str
        :sort_filename: 布尔值,True则文件夹下文件名的排序;默认False
        :add_file_name: 布尔值,True则增加字段【File_name_0】和【File_name_1】,
                        标识文件名;默认False
        :add_read_time: 布尔值,True则增加字段【File_read_Time】,标识读取时间;默认False

        return
        =======================
        :Data_result: DataFrame
        """
        Data_result = self.Origin_Read(self.File_name, unfill, TraversalExcel,
                                       changecsv, mydtype, sort_filename,
                                       add_file_name, add_read_time)

        if append_data:
            Data_result = self.Append_Data(Data_result, unfill)
        else:
            pass

        return(Data_result)

    @File_Read_Decorator
    def Origin_Read(self, File_name, unfill, TraversalExcel,
                    changecsv, mydtype, sort_filename,
                    add_file_name, add_read_time):
        """
        读取数据源
        参数
        =====================================================================
        :File_name: 当前工作环境下的文件路径
        :unfill: 布尔值,False将空值替换'';True则不替换
        :TraversalExcel: 布尔值,True则遍历工作簿下的每个非空工作表
        :changecsv: 布尔值,调用微软Excel转csv,非常规方法,尽量不使用
        :mydtype: 字符格式
        :sort_filename: 布尔值,True则文件夹下文件名的排序
        :add_file_name: 布尔值,True则增加字段【File_name_0】和【File_name_1】,
                        标识文件名
        :add_read_time: 布尔值,True则增加字段【File_read_Time】,标识读取时间

        return
        =====================
        :Data_result: 生成器
        """
        options = {
            u'unfill': unfill,
            u'TraversalExcel': TraversalExcel,
            u'changecsv': changecsv,
            u'mydtype': mydtype,
            u'sort_filename': sort_filename,
            u'add_file_name': add_file_name,
            u'add_read_time': add_read_time}

        Name_origins = os.listdir(os.getcwd() + os.sep + File_name)

        def f(x): return(x != u'Thumbs.db' and not x.startswith('~$'))
        Name_origins = list(filter(f, Name_origins))
        if sort_filename:
            Name_origins = sorted(Name_origins)
            # six.print_(Name_origins)

        Dict_method = {u'csv': self.Read_CSV, u'txt': self.Read_TXT,
                       u'xls': self.Read_Excel, u'xlsx': self.Read_Excel,
                       u'xlsb': self.Read_Excel, u'xlsm': self.Read_Excel}

        if bool(Name_origins):
            for Name_origin in Name_origins:
                Dir_origin = os.getcwd() + os.sep + File_name + \
                    os.sep + Name_origin

                Format_file = re.findall('\.(\d*\w*\d*)$', Name_origin)[0]
                Method_select = Dict_method.get(Format_file)
                if bool(Method_select):
                    Data_result = Method_select(Dir_origin,
                                                Name_origin,
                                                options)
                    while True:
                        try:
                            yield next(Data_result)
                        except StopIteration:
                            break
                else:
                    six.print_(Dir_origin, 'cannot be read error')

    def Read_CSV(self, Dir_origin, Name_origin, options):
        """
        csv读取
        参数
        ===================================
        :Dir_origin: 文件读取路径
        :Name_origin: 文件名
        :options: 选项,字典形式,关键key如下
        -----------------------------------------------------------------------
        :unfill: 布尔值,False将空值替换'';True则不替换
        :mydtype: 字符格式
        :add_file_name: 布尔值,True则增加字段【File_name_0】和【File_name_1】,
                        标识文件名
        :add_read_time: 布尔值,True则增加字段【File_read_Time】,标识读取时间
        -----------------------------------------------------------------------

        return
        ====================
        :Data_result: 生成器
        """
        try:
            Open_origin = open(Dir_origin, encoding='gb18030')
            Data_result = pd.read_csv(Open_origin,
                                      low_memory=False,
                                      keep_default_na=options.get(u'unfill'),
                                      dtype=options.get('mydtype'))
        except UnicodeDecodeError:
            Open_origin = open(Dir_origin, encoding='utf8')
            Data_result = pd.read_csv(Open_origin,
                                      low_memory=False,
                                      keep_default_na=options.get(u'unfill'),
                                      dtype=options.get('mydtype'))
        finally:
            Open_origin.close()

        if options.get(u'add_file_name'):
            Data_result = Data_result.assign(File_name_0=Name_origin)
            Data_result = Data_result.assign(File_name_1='')

        if options.get(u'add_read_time'):
            Data_result = Data_result.assign(File_read_Time=str(time.time()))

        yield Data_result

    def Read_TXT(self, Dir_origin, Name_origin, options):
        """
        txt读取
        参数
        ===================================
        :Dir_origin: 文件读取路径
        :Name_origin: 文件名
        :options: 选项,字典形式,关键key如下
        -----------------------------------------------------------------------
        :unfill: 布尔值,False将空值替换'';True则不替换
        :mydtype: 字符格式
        :add_file_name: 布尔值,True则增加字段【File_name_0】和【File_name_1】,
                        标识文件名
        :add_read_time: 布尔值,True则增加字段【File_read_Time】,标识读取时间
        -----------------------------------------------------------------------
        return
        ====================
        :Data_result: 生成器
        """
        try:
            Open_origin = open(Dir_origin, encoding='gb18030')
            Data_result = pd.read_table(Open_origin,
                                        low_memory=False,
                                        keep_default_na=options.get(u'unfill'),
                                        dtype=options.get('mydtype'))
        except UnicodeDecodeError:
            Open_origin = open(Dir_origin, encoding='utf8')
            Data_result = pd.read_table(Open_origin,
                                        low_memory=False,
                                        keep_default_na=options.get(u'unfill'),
                                        dtype=options.get('mydtype'))
        finally:
            Open_origin.close()

        if options.get(u'add_file_name'):
            Data_result = Data_result.assign(File_name_0=Name_origin)
            Data_result = Data_result.assign(File_name_1='')
        if options.get(u'add_read_time'):
            Data_result = Data_result.assign(File_read_Time=str(time.time()))

        yield Data_result

    def Read_Excel(self, Dir_origin, Name_origin, options):
        """
        Excel读取
        参数
        ==================================
        :Dir_origin: 文件读取路径
        :Name_origin: 文件名
        :options: 选项,字典形式,关键key如下
        -----------------------------------------------------------------------
        :TraversalExcel: 布尔值,True则遍历工作簿下的每个非空工作表
        :changecsv: 布尔值,若True,则调用微软Excel转csv,非常规方法,尽量不使用;
                    若False或转换失败,则直接读取xls/xlsx/xlsm/xlsb
        :unfill: 布尔值,False将空值替换'';True则不替换
        :mydtype: 字符格式
        :add_file_name: 布尔值,True则增加字段【File_name_0】和【File_name_1】,
                        标识文件名
        :add_read_time: 布尔值,True则增加字段【File_read_Time】,标识读取时间
        -----------------------------------------------------------------------
        return
        ====================
        :Data_result: 生成器
        """
        try:
            if options.get(u'changecsv'):
                normal_read_xlsx = False
                Dir_changecsv, Dir_changecsvfile = Fotmat_Change(
                    Dir_origin).Summary_Deal()

                Data_result = self.Read_CSV(Dir_changecsv,
                                            Name_origin,
                                            options)

                yield next(Data_result)
            else:
                normal_read_xlsx = True
        except Exception as e:
            normal_read_xlsx = True
            six.print_(Dir_changecsv, Dir_changecsvfile)
            six.print_('Error_Reason:【', e, '】', sep='')
            six.print_(u'Warning:【Changecsv_is_filed,Excel is being read】')
        finally:
            try:
                os.remove(Dir_changecsv)
                os.rmdir(Dir_changecsvfile)
            except BaseException:
                pass

        if normal_read_xlsx:
            Data_sheetnames = pd.ExcelFile(Dir_origin).sheet_names

            if not options.get(u'TraversalExcel'):
                Data_sheetnames = [Data_sheetnames[0]]

            for Data_sheetname in Data_sheetnames:
                Data_result = pd.read_excel(
                    Dir_origin,
                    sheetname=Data_sheetname,
                    keep_default_na=options.get(u'unfill'),
                    dtype=options.get(u'mydtype'))

                if options.get(u'add_file_name'):
                    Data_result = Data_result.assign(File_name_0=Name_origin)
                    Data_result = Data_result.assign(
                        File_name_1=Data_sheetname)
                if options.get(u'add_read_time'):
                    Data_result = Data_result.assign(
                        File_read_Time=str(time.time()))

                yield Data_result

    @File_Read_Decorator
    def Append_Data(self, Data, unfill):
        """
        合并数据
        参数
        =============
        :Data: 生成器

        return
        =======================
        :Data_result: DataFrame
        """
        Data_result = next(Data)
        while True:
            try:
                Data_result = Data_result.append(next(Data),
                                                 ignore_index=True)
            except StopIteration:
                break

        if not unfill:
            Data_result.fillna('', inplace=True)

        return(Data_result)

#Data_result = File_Read_Traversal_Append(u'数据源').Summary_File_Read(TraversalExcel=True,append_data=True,sort_filename=True,add_file_name=True,add_read_time=True)
# next(Data_result)
#Data_result.to_excel(u'e:\\data\\5.xlsx', encoding='gb18030', index=False)
