# -*- coding: utf-8 -*-
"""
Created on Tue May 23 23:28:27 2023

@author: muhammad.naveed
"""

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import Project_function

chrome_options=Options()
#chrome_options.add_argument("--headless")


chromedriver_path=r"C:\\Users\\muhammad.naveed\\Downloads\\chromedriver_win32\\chromedriver.exe"
website='https://zh.chregister.ch/cr-portal/suche/suche.xhtml'

service=Service(executable_path=chromedriver_path)

driver=webdriver.Chrome(service=service, options=chrome_options)

driver.get(website)

driver.implicitly_wait(10)

cookie_button=driver.find_element("xpath",'//*[@class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only btn btn-default"]')
cookie_button.click()
driver.implicitly_wait(2)


additional_search_button=driver.find_element("xpath",'//*[@class="ui-accordion-header ui-helper-reset ui-state-default ui-corner-all"]')
additional_search_button.click()
driver.implicitly_wait(2)


city=driver.find_element("xpath",'//*[@id="idSucheForm:panel:idSitz_input"]').send_keys("Wallisellen")
driver.implicitly_wait(2)

city_button=driver.find_element("xpath",'//*[@id="idSucheForm:panel:idSitz_panel"]/ul/li')
city_button.click()
driver.implicitly_wait(10)



values=[]

x=0

for x in range(6): 
    
    x=x+1
    
    select_button=driver.find_element("xpath",'//*[@id="idSucheForm:panel:idRechtsform"]/div[3]')
    select_button.click()
    driver.implicitly_wait(2)

    select_menu_button=driver.find_element("xpath",'//*[@id="idSucheForm:panel:idRechtsform_'+str(x)+'"]')
    select_menu_button.click()
    driver.implicitly_wait(2)
    
    search_button=driver.find_element("xpath",'//*[@class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-left btn btn-primary"]')
    search_button.click()

    time.sleep(5)
    
    
    while True:
        wait=WebDriverWait(driver, 10)
        elements=wait.until(EC.visibility_of_all_elements_located((By.XPATH, '//*[@id="idSucheForm:resultTable_data"]/tr/td[1]')))
    
    #print (elements.text)
    
        for element in elements:
            value=element.text
            values.append(value)
        
        
        next_button=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Next Page"]')))
        
        tabindex=next_button.get_attribute("tabindex")
        
        if tabindex=="-1":
            break
        
        next_button.click()
        time.sleep(5)
        
    driver.execute_script("window.scrollTo(0, 0);")
        
    
        
driver.quit()

print (values, len(values))

#Exporting Links to Excel file

df = pd.DataFrame(values, columns=['UID'])
Project_function.Convert_to_Excel(df, 'UID_list')
        