# -*- coding: utf-8 -*-
"""
Created on Wed Aug 12 16:28:27 2020

@author: PRANAVSAI
"""

# -*- coding: utf-8 -*-
"""
Created on Wed May 27 21:21:21 2020

@author: PRANAVSAI
"""


import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


dirListing = os.listdir(r'D:\Actuary\Research\!!! Life Data\Life council\pdfs')
pdfs = []
for item in dirListing:
    if ".pdf" in item:
        pdfs.append(item)






#use the following code if the link is an .aspx file
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", 'D:\Actuary\Research\!!! Life Data\Life council\pdfs')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")

driver = webdriver.Firefox(firefox_profile=profile,executable_path=r'C:\Program Files\Geckodriver\geckodriver.exe')



k=7
for i in range(len(pdfs)):    
    item = pdfs[k]
    driver.get("https://www.pdftoexcel.com/")
    
    w = WebDriverWait(driver, 40)
    w.until(EC.presence_of_element_located((By.NAME,'Filedata')))
    inputElement = driver.find_element_by_name("Filedata")
    inputElement.send_keys("D:\Actuary\Research\!!! Life Data\Life council\pdfs\\"+item)
    element = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Free")))
    element.click()
    k=k+1
    print(k)




    





















