from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc
import pandas as pd
from datetime import date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from pathlib import Path
import openpyxl 
import numpy as np
import jdatetime
import jdatetime
from datetime import datetime



error_counter = 0

class DumpToSQL:
    
    def check(self,df,col):
        if col in df:
            return True
        else:
            return False
    
    def get_update_date(self):
        x = jdatetime.date.today()
        self.update_date = str(x.year) + '/' + str(x.month) + '/' + str(x.day)
        
        return self.update_date
        
    def connect_to_sql(self,sql_query,df_values='default'):
        global error_counter
        
        def retry():
                
                server = '.'
                database = 'testdb'
                username = 'sa'
                password = '14579Ali.'
                cnxn = pyodbc.connect(
                    'DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
                cursor = cnxn.cursor()
                if df_values=='default':
                    cursor.execute(sql_query)
                    cursor.execute('commit')
                    print('done')
                    
                else:
                    cursor.executemany(sql_query, df_values)
                    cursor.execute('commit')
                    print('batch insert done')
                cnxn.close()
                    
        try:
            retry()
            
        except Exception as e:
            if error_counter < 6:
                error_counter+=1
                print (e)
                time.sleep(3)
                print('trying again')
                time.sleep(4)
                retry()


    
    def input_info(self):
    
        global input_error_counter
        try:
            print('table names:\n[testdb].[dbo].[tblGhateeSazi] = 1\n[testdb].[dbo].[tblTashkhisSaderShode] = 2\n[testdb].[dbo].[tblTashkhisEblaghShode] = 3\n[testdb].[dbo].[tblGhateeSaderShode] = 4\n[testdb].[dbo].[tblGhateeEblaghShode] = 5\n\n') 
                
            table = int(input("Enter table: "))
            
            if (table == 1):
                table = '[testdb].[dbo].[tblGhateeSazi]'
                report_type = 'ezhar'
            elif (table == 2):
                table = '[testdb].[dbo].[tblTashkhisSaderShode]'
                report_type = 'tashkhis_sader_shode'
            elif (table == 3):
                table = '[testdb].[dbo].[tblTashkhisEblaghShode]'
                report_type = 'tashkhis_eblagh_shode'
            elif (table == 4):
                table = '[testdb].[dbo].[tblGhateeSaderShode]'
                report_type = 'ghatee_sader_shode'
            elif (table == 5):
                table = '[testdb].[dbo].[tblGhateeEblaghShode]'
                report_type = 'ghatee_eblagh_shode'
                
         
            return table, report_type
        
        except:
            print("Please Enter integer values")
            if input_error_counter < 6:
                input_error_counter+=1
                self.input_info()

    

    def create_sql_table(self, table):
        print(table)        
        if table == '[testdb].[dbo].[tblGhateeSazi]':
            sql_query = """\
            IF Object_ID('tblGhateeSazi') IS NULL
            
            CREATE TABLE [testdb].[dbo].[tblGhateeSazi]
            (
             [ID] [int] IDENTITY(1,1) NOT NULL,
                                         [کد اداره] NVARCHAR(MAX) NULL,        
             [نام اداره] NVARCHAR(MAX) NULL,        
             [سال عملکرد] NVARCHAR(MAX) NULL,        
             ['شناسه ملی / کد ملی (TIN)] NVARCHAR(MAX) NULL,       
             [کد رهگیری ثبت نام] NVARCHAR(MAX) NULL,       
             [نوع مودی] NVARCHAR(MAX) NULL,       
             [نام مودی] NVARCHAR(MAX) NULL,       
             [کدپستی مودی] NVARCHAR(MAX) NULL,       
             [شناسه اظهارنامه] NVARCHAR(MAX) NULL,       
               [کانال تسلیم اظهارنامه] NVARCHAR(MAX) NULL,       
              [دوره] NVARCHAR(MAX) NULL,       
              [منبع مالیاتی] NVARCHAR(MAX) NULL,       
              [تاریخ تسلیم اظهارنامه] NVARCHAR(MAX) NULL,       
              [درآمد ابرازی] NVARCHAR(MAX) NULL,       
              [مالیات ابرازی] NVARCHAR(MAX) NULL,       
              [عوارض ابرازی] NVARCHAR(MAX) NULL,       
              [فروش ابرازی] NVARCHAR(MAX) NULL,       
              [اعتبار ابرازی] NVARCHAR(MAX) NULL,       
              [منبع ورود اظهارنامه] NVARCHAR(MAX) NULL,       
             [کد مرسوله] NVARCHAR(MAX) NULL,       
             [نوع ریسک اظهارنامه] NVARCHAR(MAX) NULL,       
             [دارای برگ تشخیص] NVARCHAR(MAX) NULL,       
             [دارای برگ قطعی] NVARCHAR(MAX) NULL,       
             [آغاز سال مالی] NVARCHAR(MAX) NULL,       
             [پایان سال مالی] NVARCHAR(MAX) NULL,       
             [میان سالی] NVARCHAR(MAX) NULL,       
             [کد رهگیری اظهارنامه] NVARCHAR(MAX) NULL,       
             [تاریخ بروزرسانی] NVARCHAR(MAX) NULL,       
       PRIMARY KEY (ID)
             )
            """
        elif table =='[testdb].[dbo].[tblTashkhisSaderShode]':
             sql_query = """\
        IF Object_ID('tblTashkhisSaderShode') IS NULL
        
        CREATE TABLE [testdb].[dbo].[tblTashkhisSaderShode]
        (
         [ID] [int] IDENTITY(1,1) NOT NULL,
                                     [کد اداره] NVARCHAR(MAX) NULL,        
         [نام اداره] NVARCHAR(MAX) NULL,        
         [سال عملکرد] NVARCHAR(MAX) NULL,        
         [شناسه ملی / کد ملی (TIN)] NVARCHAR(MAX) NULL,       
         [نوع مودی] NVARCHAR(MAX) NULL,       
         [نام مودی] NVARCHAR(MAX) NULL,       
         [کدپستی مودی] NVARCHAR(MAX) NULL,       
         [شناسه اظهارنامه] NVARCHAR(MAX) NULL,       
          [دوره] NVARCHAR(MAX) NULL,       
          [منبع مالیاتی] NVARCHAR(MAX) NULL,       
          [نام حسابرس اصلی] NVARCHAR(MAX) NULL,       
          [کد ملی حسابرس اصلی] NVARCHAR(MAX) NULL,       
          [تاریخ صدور برگه تشخیص] NVARCHAR(MAX) NULL,       
          [شماره برگه تشخیص] NVARCHAR(MAX) NULL,       
          [وضعیت برگ تشخیص] NVARCHAR(MAX) NULL,       
          [مالیات تشخیص] NVARCHAR(MAX) NULL,       
          [درآمد تشخیص] NVARCHAR(MAX) NULL,       
         [عوارض تشخیص] NVARCHAR(MAX) NULL,       
         [فروش تشخیص] NVARCHAR(MAX) NULL,       
         [اعتبار تشخیص] NVARCHAR(MAX) NULL,       
         [برگ مطالبه جرایم موضوع ماده 169 ق.م.م] NVARCHAR(MAX) NULL,       
         [نوع ریسک اظهارنامه] NVARCHAR(MAX) NULL,       
         [تاریخ بروزرسانی] NVARCHAR(MAX) NULL,       
        PRIMARY KEY (ID)
         )
        """
        
        elif table == '[testdb].[dbo].[tblTashkhisEblaghShode]':
            sql_query = """\
            IF Object_ID('tblTashkhisEblaghShode') IS NULL
            
            CREATE TABLE [testdb].[dbo].[tblTashkhisEblaghShode]
            (
             [ID] [int] IDENTITY(1,1) NOT NULL,
                                         [کد اداره] NVARCHAR(MAX) NULL,        
             [نام اداره] NVARCHAR(MAX) NULL,        
             [سال عملکرد] NVARCHAR(MAX) NULL,        
             [شناسه ملی / کد ملی (TIN)] NVARCHAR(MAX) NULL,       
             [نوع مودی] NVARCHAR(MAX) NULL,       
             [نام مودی] NVARCHAR(MAX) NULL,       
             [کدپستی مودی] NVARCHAR(MAX) NULL,       
             [شناسه اظهارنامه] NVARCHAR(MAX) NULL,       
              [دوره] NVARCHAR(MAX) NULL,       
              [منبع مالیاتی] NVARCHAR(MAX) NULL,       
              [نام حسابرس اصلی] NVARCHAR(MAX) NULL,       
              [کد ملی حسابرس اصلی] NVARCHAR(MAX) NULL,       
              [تاریخ صدور برگه تشخیص] NVARCHAR(MAX) NULL,       
              [شماره برگه تشخیص] NVARCHAR(MAX) NULL,       
              [وضعیت برگ تشخیص] NVARCHAR(MAX) NULL,       
              [مالیات تشخیص] NVARCHAR(MAX) NULL,       
              [درآمد تشخیص] NVARCHAR(MAX) NULL,       
             [عوارض تشخیص] NVARCHAR(MAX) NULL,       
             [فروش تشخیص] NVARCHAR(MAX) NULL,       
             [اعتبار تشخیص] NVARCHAR(MAX) NULL,       
        [ابلاغ الکترونیک] NVARCHAR(MAX) NULL,       
         [تاریخ ابلاغ] NVARCHAR(MAX) NULL,       
 [مامور ابلاغ] NVARCHAR(MAX) NULL,       
     [نوع ابلاغ] NVARCHAR(MAX) NULL,       
             [برگ مطالبه جرایم موضوع ماده 169 ق.م.م] NVARCHAR(MAX) NULL,       
             [نوع ریسک اظهارنامه] NVARCHAR(MAX) NULL,       
             [تاریخ بروزرسانی] NVARCHAR(MAX) NULL,       
        PRIMARY KEY (ID)
             )
            """
            
        elif table == '[testdb].[dbo].[tblGhateeSaderShode]':
            sql_query = """\
            IF Object_ID('tblGhateeSaderShode') IS NULL
            
            CREATE TABLE [testdb].[dbo].[tblGhateeSaderShode]
            (
             [ID] [int] IDENTITY(1,1) NOT NULL,
                                         [کد اداره] NVARCHAR(MAX) NULL,        
             [نام اداره] NVARCHAR(MAX) NULL,        
             [شناسه ملی / کد ملی (TIN)] NVARCHAR(MAX) NULL,       
             [نوع مودی] NVARCHAR(MAX) NULL,       
             [نام مودی] NVARCHAR(MAX) NULL,       
             [کدپستی مودی] NVARCHAR(MAX) NULL,       
             [سال عملکرد] NVARCHAR(MAX) NULL,       
              [منبع مالیاتی] NVARCHAR(MAX) NULL,       
                            [شناسه اظهارنامه] NVARCHAR(MAX) NULL,       
                              [دوره] NVARCHAR(MAX) NULL,       
              [شماره برگه قطعی] NVARCHAR(MAX) NULL,       
              [وضعیت برگه قطعی] NVARCHAR(MAX) NULL,       
[تاریخ برگ قطعی] NVARCHAR(MAX) NULL,       
[مالیات قطعی] NVARCHAR(MAX) NULL,       
[درآمد قطعی] NVARCHAR(MAX) NULL,       
[عوارض قطعی] NVARCHAR(MAX) NULL,       
[فروش قطعی] NVARCHAR(MAX) NULL,       
[اعتبار قطعی] NVARCHAR(MAX) NULL,       
[صادر کننده] NVARCHAR(MAX) NULL,       
[پرداخت] NVARCHAR(MAX) NULL,       
[مانده بدهی] NVARCHAR(MAX) NULL,       
[شماره برگ تشخیص] NVARCHAR(MAX) NULL,       
              [منبع قطعی] NVARCHAR(MAX) NULL,       
              [مالیات] NVARCHAR(MAX) NULL,       
              [نوع ریسک اظهارنامه] NVARCHAR(MAX) NULL,       
             [تاریخ بروزرسانی] NVARCHAR(MAX) NULL,       
       PRIMARY KEY (ID)
             )
            """
            
        elif table == '[testdb].[dbo].[tblGhateeEblaghShode]':
            sql_query = """\
            IF Object_ID('tblGhateeEblaghShode') IS NULL
            
            CREATE TABLE [testdb].[dbo].[tblGhateeEblaghShode]
            (
             [ID] [int] IDENTITY(1,1) NOT NULL,
                                         [کد اداره] NVARCHAR(MAX) NULL,        
             [نام اداره] NVARCHAR(MAX) NULL,        
             [شناسه ملی / کد ملی (TIN)] NVARCHAR(MAX) NULL,       
             [نوع مودی] NVARCHAR(MAX) NULL,       
             [نام مودی] NVARCHAR(MAX) NULL,       
             [کدپستی مودی] NVARCHAR(MAX) NULL,       
                          [سال عملکرد] NVARCHAR(MAX) NULL,       
            [منبع مالیاتی] NVARCHAR(MAX) NULL,       
                            [شناسه اظهارنامه] NVARCHAR(MAX) NULL,       
                              [دوره] NVARCHAR(MAX) NULL,       
              [شماره برگه قطعی] NVARCHAR(MAX) NULL,       
              [وضعیت برگه قطعی] NVARCHAR(MAX) NULL,       
[تاریخ برگ قطعی] NVARCHAR(MAX) NULL,       
[مالیات قطعی] NVARCHAR(MAX) NULL,       
[درآمد قطعی] NVARCHAR(MAX) NULL,       
[عوارض قطعی] NVARCHAR(MAX) NULL,       
[فروش قطعی] NVARCHAR(MAX) NULL,       
[اعتبار قطعی] NVARCHAR(MAX) NULL,       
[صادر کننده] NVARCHAR(MAX) NULL,       
[پرداخت] NVARCHAR(MAX) NULL,       
[مانده بدهی] NVARCHAR(MAX) NULL,       
[شماره برگ تشخیص] NVARCHAR(MAX) NULL,       
              [منبع قطعی] NVARCHAR(MAX) NULL,       
                            [ابلاغ الکترونیک] NVARCHAR(MAX) NULL,       
              [تاریخ ابلاغ برگ قطعی] NVARCHAR(MAX) NULL,       
              [مالیات] NVARCHAR(MAX) NULL,       
              [مامور ابلاغ] NVARCHAR(MAX) NULL,       
              [نوع ابلاغ] NVARCHAR(MAX) NULL,       
              [نوع ریسک اظهارنامه] NVARCHAR(MAX) NULL,       
             [تاریخ بروزرسانی] NVARCHAR(MAX) NULL,       
       PRIMARY KEY (ID)
             )
            """
        
        self.connect_to_sql(sql_query)
        
################################################################################################################################

    def insert_into(self,table, df_values):
        if table == '[testdb].[dbo].[tblGhateeSazi]':
            
            sql_insert = """\
                    
        INSERT INTO [testdb].[dbo].[tblGhateeSazi]
        (
             [کد اداره],        
             [نام اداره],        
             [سال عملکرد],        
             [شناسه ملی / کد ملی (TIN)],        
             [کد رهگیری ثبت نام],        
             [نوع مودی],        
             [نام مودی],        
             [کدپستی مودی],        
             [شناسه اظهارنامه],        
             [کانال تسلیم اظهارنامه],        
             [دوره],        
             [منبع مالیاتی],        
             [تاریخ تسلیم اظهارنامه],        
             [درآمد ابرازی],        
             [مالیات ابرازی],        
             [عوارض ابرازی],        
             [فروش ابرازی],        
             [اعتبار ابرازی],        
             [منبع ورود اظهارنامه],        
             [کد مرسوله],        
             [نوع ریسک اظهارنامه],        
             [دارای برگ تشخیص],        
             [دارای برگ قطعی],        
             [آغاز سال مالی],        
             [پایان سال مالی],        
             [میان سالی],        
             [کد رهگیری اظهارنامه],       
              [تاریخ بروزرسانی]        
         )
        
        VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """
        elif table =='[testdb].[dbo].[tblTashkhisSaderShode]':
            
             sql_insert = """\
                    
        INSERT INTO [testdb].[dbo].[tblTashkhisSaderShode]
        (
              [کد اداره],        
             [نام اداره],        
             [سال عملکرد],        
             [شناسه ملی / کد ملی (TIN)],        
             [نوع مودی],        
             [نام مودی],        
             [کدپستی مودی],        
             [شناسه اظهارنامه],        
             [دوره],        
             [منبع مالیاتی],        
              [نام حسابرس اصلی],        
            [کد ملی حسابرس اصلی],        
            [تاریخ صدور برگه تشخیص],        
             [شماره برگه تشخیص],        
              [وضعیت برگ تشخیص],        
             [مالیات تشخیص],        
            [درآمد تشخیص],        
            [عوارض تشخیص],        
             [فروش تشخیص],        
            [اعتبار تشخیص],        
              [برگ مطالبه جرایم موضوع ماده 169 ق.م.م],        
           [نوع ریسک اظهارنامه],        
              [تاریخ بروزرسانی]         
         )
        
        VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? )
        """
        
        elif table == '[testdb].[dbo].[tblTashkhisEblaghShode]':
            
            sql_insert = """\
                    
        INSERT INTO [testdb].[dbo].[tblTashkhisEblaghShode]
        (
             [کد اداره],        
             [نام اداره],        
             [سال عملکرد],        
             [شناسه ملی / کد ملی (TIN)],        
             [نوع مودی],        
             [نام مودی],        
             [کدپستی مودی],        
             [شناسه اظهارنامه],        
             [دوره],        
             [منبع مالیاتی],        
            [نام حسابرس اصلی],        
             [کد ملی حسابرس اصلی],        
             [تاریخ صدور برگه تشخیص],        
             [شماره برگه تشخیص],        
           [وضعیت برگ تشخیص],        
            [مالیات تشخیص],        
             [درآمد تشخیص],        
           [عوارض تشخیص],        
            [فروش تشخیص],        
           [اعتبار تشخیص],        
           [ابلاغ الکترونیک],        
            [تاریخ ابلاغ],        
            [مامور ابلاغ],        
            [نوع ابلاغ],        
            [برگ مطالبه جرایم موضوع ماده 169 ق.م.م],       
                       [نوع ریسک اظهارنامه],        
              [تاریخ بروزرسانی]        
         )
        
        VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """
        
        elif table == '[testdb].[dbo].[tblGhateeSaderShode]':
            
            sql_insert = """\
                    
        INSERT INTO [testdb].[dbo].[tblGhateeSaderShode]
        (
             [کد اداره],        
             [نام اداره],        
             [شناسه ملی / کد ملی (TIN)],        
             [نوع مودی],        
             [نام مودی],        
             [کدپستی مودی],        
                          [سال عملکرد],        
                                       [منبع مالیاتی],        
             [شناسه اظهارنامه],        
             [دوره],        
              [شماره برگه قطعی],        
          [وضعیت برگه قطعی],        
         [تاریخ برگ قطعی],        
            [مالیات قطعی],        
             [درآمد قطعی],        
            [عوارض قطعی],        
           [فروش قطعی],        
          [اعتبار قطعی],        
                    [صادر کننده],        
            پرداخت,        
            [مانده بدهی],        
          [شماره برگ تشخیص],        
                  [منبع قطعی],        
        مالیات,        
           [نوع ریسک اظهارنامه],        
              [تاریخ بروزرسانی]         
         )
        
        VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """
            
        elif table == '[testdb].[dbo].[tblGhateeEblaghShode]':
            sql_insert = """\
                    
        INSERT INTO [testdb].[dbo].[tblGhateeEblaghShode]
        (
              [کد اداره],        
             [نام اداره],        
             [شناسه ملی / کد ملی (TIN)],        
             [نوع مودی],        
             [نام مودی],        
             [کدپستی مودی],        
                          [سال عملکرد],        
                                       [منبع مالیاتی],        
             [شناسه اظهارنامه],        
             [دوره],        
              [شماره برگه قطعی],        
          [وضعیت برگه قطعی],        
         [تاریخ برگ قطعی],        
            [مالیات قطعی],        
             [درآمد قطعی],        
            [عوارض قطعی],        
           [فروش قطعی],        
          [اعتبار قطعی],        
                    [صادر کننده],        
            پرداخت,        
            [مانده بدهی],        
          [شماره برگ تشخیص],        
                  [منبع قطعی],        
                                  [ابلاغ الکترونیک],        
                  [تاریخ ابلاغ برگ قطعی],        
        مالیات,        
                       [مامور ابلاغ],        
                [نوع ابلاغ],        
           [نوع ریسک اظهارنامه],        
              [تاریخ بروزرسانی]         
         )
        
        VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """

        
        self.connect_to_sql(sql_insert, df_values)
        

    def dump_to_sql(self,years):
        
        table, report_type = self.input_info()

        merge_excels=[]
        
        for year in years:
            
            path =r'J:\ezhar-temp\%s\%s' % (year, report_type)
            
            excel_files = glob.glob(os.path.join(path, "*.xlsx"))
            
            
            for f in excel_files:
                print('opening %s for saving\n' %f)
                irismash = xw.Book(f)
                irismash.save()
                irismash.app.quit()
                # read the csv file
                df = pd.read_excel(f)
                merge_excels.append(df)
               
        final_df_all_fine_grained = pd.concat(merge_excels)
        final_df_all_fine_grained = final_df_all_fine_grained.fillna(value=0)
        final_df_all_fine_grained['تاریخ بروزرسانی']=self.get_update_date()
        final_df_all_fine_grained['شناسه ملی / کد ملی (TIN)'] = final_df_all_fine_grained['شناسه ملی / کد ملی (TIN)'].astype(str)
        final_df_all_fine_grained['شناسه اظهارنامه']= final_df_all_fine_grained['شناسه اظهارنامه'].astype(str)


        if (report_type=='ezhar'):
            final_df_all_fine_grained['کد رهگیری ثبت نام']= final_df_all_fine_grained['کد رهگیری ثبت نام'].astype(str)
            final_df_all_fine_grained['عوارض ابرازی']= final_df_all_fine_grained['عوارض ابرازی'].astype(np.int64)
            final_df_all_fine_grained['کد رهگیری اظهارنامه']= final_df_all_fine_grained['کد رهگیری اظهارنامه'].astype(str)
            final_df_all_fine_grained['فروش ابرازی']= final_df_all_fine_grained['فروش ابرازی'].astype(np.int64)
            final_df_all_fine_grained['اعتبار ابرازی']= final_df_all_fine_grained['اعتبار ابرازی'].astype(np.int64)
            
        elif (report_type=='tashkhis_sader_shode'   or report_type=='tashkhis_eblagh_shode' or report_type == 'ghatee_sader_shode' or report_type == 'ghatee_eblagh_shode'):
            if (self.check(final_df_all_fine_grained, "کد ملی حسابرس اصلی")):
                final_df_all_fine_grained['کد ملی حسابرس اصلی']= final_df_all_fine_grained['کد ملی حسابرس اصلی'].astype(np.int64)
                
            if (self.check(final_df_all_fine_grained, "شماره برگه تشخیص")):
                final_df_all_fine_grained['شماره برگه تشخیص']= final_df_all_fine_grained['شماره برگه تشخیص'].astype(str)
               
            if (self.check(final_df_all_fine_grained, "مالیات تشخیص")):
                final_df_all_fine_grained['مالیات تشخیص']= final_df_all_fine_grained['مالیات تشخیص'].astype(np.int64)
            
            if (self.check(final_df_all_fine_grained, "فروش تشخیص")):
                final_df_all_fine_grained['فروش تشخیص']= final_df_all_fine_grained['فروش تشخیص'].astype(np.int64)
            
            if (self.check(final_df_all_fine_grained, "درآمد تشخیص")):
                final_df_all_fine_grained['درآمد تشخیص']= final_df_all_fine_grained['درآمد تشخیص'].astype(np.int64)
                
            if (self.check(final_df_all_fine_grained, "عوارض تشخیص")):
                final_df_all_fine_grained['عوارض تشخیص']= final_df_all_fine_grained['عوارض تشخیص'].astype(np.int64)
                
            if (self.check(final_df_all_fine_grained, "درآمد قطعی")):
                final_df_all_fine_grained['درآمد قطعی']= final_df_all_fine_grained['درآمد قطعی'].astype(np.int64)
                
            if (self.check(final_df_all_fine_grained, "عوارض قطعی")):
                final_df_all_fine_grained['عوارض قطعی']= final_df_all_fine_grained['عوارض قطعی'].astype(np.int64)
                
            if (self.check(final_df_all_fine_grained, "اعتبار قطعی")):
                final_df_all_fine_grained['اعتبار قطعی']= final_df_all_fine_grained['اعتبار قطعی'].astype(np.int64)
                
            if (self.check(final_df_all_fine_grained, "پرداخت")):
                final_df_all_fine_grained['پرداخت']= final_df_all_fine_grained['پرداخت'].astype(np.int64)   
                
            if (self.check(final_df_all_fine_grained, "مانده بدهی")):
                final_df_all_fine_grained['مانده بدهی']= final_df_all_fine_grained['مانده بدهی'].astype(np.int64)  
                
            if (self.check(final_df_all_fine_grained, "شماره برگ تشخیص")):
                final_df_all_fine_grained['شماره برگ تشخیص']= final_df_all_fine_grained['شماره برگ تشخیص'].astype(np.int64)   
                
            if (self.check(final_df_all_fine_grained, "مالیات")):
                final_df_all_fine_grained['مالیات']= final_df_all_fine_grained['مالیات'].astype(np.int64)                   
        df_values = final_df_all_fine_grained.values.tolist()

        #self.create_sql_table(table)
        
        sql_delete = "DELETE FROM %s" %table
        self.connect_to_sql(sql_delete)
        self.insert_into(table,df_values)

    
        
        
dump = DumpToSQL()
dump.dump_to_sql([1395])