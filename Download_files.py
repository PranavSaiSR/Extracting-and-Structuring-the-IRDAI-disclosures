# -*- coding: utf-8 -*-
"""
Created on Wed May 27 21:21:21 2020

@author: PRANAVSAI
"""

from bs4 import BeautifulSoup
import urllib.request
from urllib.parse import urlparse
import re
import requests
import os
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


mainlink="http://www.itireinsurance.com/PublicDisclosure.aspx"
sublink="NL"


##The below code extracts the links of all the pdf files available in a website##
class AppURLopener(urllib.request.FancyURLopener):
    version = "Mozilla/5.0"
opener = AppURLopener()
response = opener.open(mainlink)
soup = BeautifulSoup(response, "lxml")

links = []
for link in soup.findAll('a', attrs={'href': re.compile(sublink)}):
    links.append(link.get('href'))
    
ilinks = []

i=19
for i in range(len(links)):
    
    u=links[i]
    class AppURLopener(urllib.request.FancyURLopener):
        version = "Mozilla/5.0"
    opener = AppURLopener()
    response = opener.open("https://www.iffcotokio.co.in/"+u)
    soup = BeautifulSoup(response, "lxml")

    
    for link in soup.findAll('a', attrs={'href': re.compile("NL")}):
        ilinks.append(link.get('href'))
    print(i)
    
    
iilinks=[]
hg
        



#use the following code if the link is an .aspx file
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", 'C:\data of comp\Sent already\Oriental -1\pdfs')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")

driver = webdriver.Firefox(firefox_profile=profile,executable_path=r'C:\Program Files\Geckodriver\geckodriver.exe')




j=1
for i in range((len(links))):
    driver.get("https://www.iffcotokio.co.in/"+links[j] )

    w = WebDriverWait(driver, 40)
    w.until(EC.presence_of_element_located((By.ID,'download')))
    element = driver.find_element_by_id("download")
    element.click()
    j=j+1
    print(j)

  

pdfkit.from_url("https://www.gicofindia.com/"+ilinks[i], 'out.pdf')
pdf = weasyprint.HTML("https://www.gicofindia.com/"+ilinks[i]).write_pdf()

      

#use this code if the link is a pdf
    
i=2
j=1  
k=1 


    
 
dirListing = os.listdir(r'C:\data of comp\Sent already\Oriental -1\pdfs')
pdfs = []
for item in dirListing:
    if ".pdf" in item:
        pdfs.append(item)
k=i+1   

driver = webdriver.Firefox(executable_path=r'C:\Program Files\Geckodriver\geckodriver.exe') 
driver = webdriver.Chrome(executable_path=r'C:\Program Files\Chromedriver\chromedriver.exe') 

k=0

for i in range(len(pdfs)):    
    item = pdfs[k]
    driver.get("https://www.pdftoexcel.com/")
    
    w = WebDriverWait(driver, 40)
    w.until(EC.presence_of_element_located((By.NAME,'Filedata')))
    inputElement = driver.find_element_by_name("Filedata")
    inputElement.send_keys("C:\data of comp\Sent already\Oriental -1\pdfs\\"+item)
    element = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Free")))
    element.click()
    k=k+1
    print(k)

driver.quit()

j=1

for i in range(len(links)):
    
    url = "http://www.itireinsurance.com/" + links[i]
    name = str(k)+" " + os.path.basename(urlparse(url).path)
    s = os.path.join('C:\data of comp\ITI',name)
    r = requests.get(url, allow_redirects=True)
    open(s, 'wb').write(r.content)
    





















