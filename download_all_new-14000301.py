from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os
from datetime import datetime
import xlwings as xw
import pyodbc
import pandas as pd
from datetime import date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from pathlib import Path
from gozaresh import Gozareshat




error_counter = 0
input_error_counter = 0
year = 0
report_type = 0


def input_info():
    
    global input_error_counter
    try:
        print('year types:\n1395 = 1\n1396 = 2\n1397 = 3\n1398 = 4\n1399 = 5\n\n \
              Report types:\nezhar = 1\nTashkhisSaderShode = 2\nTashkhisEblaghShode = 3\nGhateeSaderShode = 4\nGhateeEblaghShode = 5')
             
        year = int(input("Enter year: "))
        report_type = int(input("Enter report type: "))
        
        if (year == 1):
            year = 1395
        elif (year == 2):
            year = 1396
        elif (year == 3):
            year = 1397
        elif (year == 4):
            year = 1398
        elif (year == 5):
            year = 1399
        elif (year == 6):
            year = 1400
            
        if (report_type == 1):
            report_type = 'ezhar'
        elif (report_type == 2):
            report_type = 'tashkhis_sader_shode'
        elif (report_type == 3):
            report_type = 'tashkhis_eblagh_shode'
        elif (report_type == 4):
            report_type = 'ghatee_sader_shode'
        elif (report_type == 5):
            report_type = 'ghatee_eblagh_shode'
        
        return year, report_type
    
    except:
        print("Please Enter integer values")
        if input_error_counter < 6:
            input_error_counter+=1
            input_info()
            
    

def scrape():
    global error_counter
    global year
    global report_type
    
        
    first_list=[4,8]
    second_list=[9,20,21]

    try:
        if (year == 0 and report_type == 0):     
            year, report_type = input_info()
        
        if (report_type == 'ezhar'):
            td_number = 4
        elif (report_type == 'tashkhis_sader_shode'):
            td_number = 8
        elif (report_type == 'tashkhis_eblagh_shode'):
            td_number = 9
        elif (report_type == 'ghatee_sader_shode'):
            td_number = 20
        elif (report_type == 'ghatee_eblagh_shode'):
            td_number = 21
        
        path = r'J:\ezhar-temp\%s\%s' % (year, report_type) 
            
        
         # Check if Gozareshat is up to date
        excel_files = glob.glob(os.path.join(path, "*.xlsx"))
        excel_files_to_be_removed=glob.glob(os.path.join(path, "*.xlsx"))
        today = datetime.today().strftime('%Y-%m-%d')
        
        uptodate=[]
        
        if (len(excel_files) > 0):
            for f in excel_files:
                # file modification timestamp of a file
                today = datetime.today().strftime('%Y-%m-%d')
                # file modification timestamp of a file
                m_time = os.path.getmtime(f)
                # convert timestamp into DateTime object
                dt_m = datetime.fromtimestamp(m_time)
                dt_m =  dt_m.strftime('%Y-%m-%d')
                if(dt_m != today):
                    continue
                else:
                    uptodate.append(f)
                    excel_files_to_be_removed.remove(f)
            
        if(len(uptodate) == 3):
            print('All excel files are up to date')
            
            return
        
        else:    
            gozareshat = Gozareshat(path)
            
            if(len(excel_files_to_be_removed) > 0):
                print('Removing excel files..........................')
                gozareshat.remove_excel_files(path,excel_files_to_be_removed)
                print('done removing excel files..........................')
            
            # Update the reports
            print('updating excel files................................')
            gozareshat.login_sanim(path)
            gozareshat.driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/span/span').click()
            time.sleep(1)
            gozareshat.driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/button').click()            
            
            WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/ul/li[1]/div/span[1]/a')))
            gozareshat.driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/ul/li[1]/div/span[1]/a').click()
            
            time.sleep(3)
            
            WebDriverWait(gozareshat.driver, 8).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div[1]/div/div/div/div[2]/div/div/button')))
            gozareshat.driver.find_element(By.XPATH,'/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div[1]/div/div/div/div[2]/div/div/button').click()  
            
            gozareshat.driver.find_element(By.XPATH,'/html/body/div[7]/div[2]/div[1]/input').send_keys(year)
            
            WebDriverWait(gozareshat.driver, 8).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[7]/div[2]/div[1]/button')))
            gozareshat.driver.find_element(By.XPATH,'/html/body/div[7]/div[2]/div[1]/button').click()  
            
            time.sleep(3)
            
            WebDriverWait(gozareshat.driver, 8).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[7]/div[2]/div[2]/div/div[3]/ul/li')))
            gozareshat.driver.find_element(By.XPATH,'/html/body/div[7]/div[2]/div[2]/div/div[3]/ul/li').click() 
            
            #################################################################################################################################
            
            time.sleep(3)
            
            WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
            gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[%s]/a' % td_number).click()
            
            time.sleep(4)
            
            exists_in_first_list = first_list.count(td_number)

            if (exists_in_first_list):
            
                if(uptodate.count(path + '\Excel.xlsx') == 0):
                    print('updating hoghoghi')
                    WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[5]/a')))
                    gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[5]/a').click()
                    
                    time.sleep(4)
                    WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]')))
                    gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]').click()
                    
                    
                    print('*******************************************************************************************')
                    
                    i=0
                    while len(glob.glob1(path, '*.xlsx')) == 0:
                        print('waiting %s seconds for the file to be downloaded' % i)
                        i+=2
                        time.sleep(2)
                    
                    print('****************Hoghoghi done*******************************')
                    
                    gozareshat.driver.back()
                    time.sleep(10)
                    
                if(uptodate.count(path + '\Excel(1).xlsx') == 0):
                    print('updating haghighi')
                    WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
                    gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
                    
                    time.sleep(4)
                    WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]')))
                    gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]').click()
                    
                    print('*******************************************************************************************')
                    
                    i=0
                    while len(glob.glob1(path, '*.xlsx')) == 1:
                        print('waiting %s seconds for the file to be downloaded' % i)
                        i+=2
                        time.sleep(2)
                    
                    time.sleep(10)
                    
                    print('****************Haghighi done*******************************')
                    
                    gozareshat.driver.back()
                    
                if(uptodate.count(path + '\Excel(2).xlsx') == 0):
                    print('updating Arzesh Afzoode')
                    WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[8]/a')))
                    gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[8]/a').click()
                    
                    time.sleep(4)
                    WebDriverWait(gozareshat.driver, 180).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]')))
                    gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]').click()
                    
                    print('*******************************************************************************************')
                    
                    i=0
                    while len(glob.glob1(path, '*.xlsx')) == 2:
                        print('waiting %s seconds for the file to be downloaded' % i)
                        i+=2
                        time.sleep(2)
                    
                    time.sleep(10)
                    print('****************Arzesh Afzoode done*******************************')
                    
                    gozareshat.driver.back()
                
                time.sleep(180)
                
                gozareshat.driver.close()

                
            else:
                time.sleep(4)
                WebDriverWait(gozareshat.driver, 500).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]')))
                gozareshat.driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]').click()
                
                i=0
                while len(glob.glob1(path, '*.xlsx')) == 0:
                    print('waiting %s seconds for the file to be downloaded' % i)
                    i+=2
                    time.sleep(2)
                            
                gozareshat.driver.back()
                
                time.sleep(180)
                
                gozareshat.driver.close()


            
    except Exception as e:
        if error_counter < 6:
            error_counter+=1
            print (e)
            time.sleep(3)
            gozareshat.driver.close()
            print('trying again')
            time.sleep(4)
            scrape()
               
    

if __name__ == '__main__':
    
    scrape()




    
    ################################################################################################################################    
        