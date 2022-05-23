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


class Gozareshat:
    def __init__(self, pathsave):
        fp = webdriver.FirefoxProfile()
        fp.set_preference('browser.download.folderList', 2)
        fp.set_preference('browser.download.manager.showWhenStarting', False)
        fp.set_preference('browser.download.dir', pathsave)
        fp.set_preference('browser.helperApps.neverAsk.openFile',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        fp.set_preference('browser.helperApps.neverAsk.saveToDisk',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        fp.set_preference('browser.helperApps.alwaysAsk.force', False)
        fp.set_preference('browser.download.manager.alertOnEXEOpen', False)
        fp.set_preference('browser.download.manager.focusWhenStarting', False)
        fp.set_preference('browser.download.manager.useWindow', False)
        fp.set_preference('browser.download.manager.showAlertOnComplete', False)
        fp.set_preference('browser.download.manager.closeWhenDone', False)
        
        self.driver = webdriver.Firefox(fp, executable_path="H:\driver\geckodriver.exe")
        self.driver.window_handles
        self.driver.switch_to.window(self.driver.window_handles[0])    
        
        
    def login_sanim(self,pathsave):
       
        self.driver.get("https://mgmt.tax.gov.ir/ords/f?p=100:101:16540338045165:::::")
        self.driver.implicitly_wait(20)
        txtUserName = self.driver.find_element_by_id('P101_USERNAME').send_keys('1970619521')
        txtPassword = self.driver.find_element_by_id('P101_PASSWORD').send_keys('123456')
        
        self.driver.find_element(By.ID, 'B1700889564218640').click()
    

        
    
    def remove_excel_files(self,pathsave,files):
        for f in files:
            os.remove(f)
            
            
    