# -*- coding: utf-8 -*-
"""
Created on Mon May 29 18:14:17 2023

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


#reading scraped links from excel file
df=pd.read_excel(r"D:\Naveed\\Freelancing\\1. Upwork\\2. Work for Portfolio\\2. Jobs to be performed\\job06\\Final Programs\\Outputs\\UID_list.xlsx")
uids=df['UID'].tolist()


#if program is stopped in middle, this will check for last processed data to avoid repeataation.
try:
    df1=pd.read_excel(r"D:\Naveed\\Freelancing\\1. Upwork\\2. Work for Portfolio\\2. Jobs to be performed\\job06\\Final Programs\\Outputs\\Completed_UID_list.xlsx")
    completed_uids=df1['Completed_UID'].tolist()
    uids= [item for item in uids if item not in completed_uids]
    
except:
    pass
    

chrome_options=Options()

chrome_options.headless=False
#chrome_options.add_argument("--start-minimized") #to keep window minimized
#chrome_options.add_argument("--window-position=-2000,0")

chromedriver_path=r"C:\\Users\\muhammad.naveed\\Downloads\\chromedriver_win32\\chromedriver.exe"



#numbers=['CHE-493.079.735', 'CHE-290.613.181', 'CHE-112.093.055', 'CHE-227.142.651', 'CHE-160.078.638', 'CHE-113.382.042']#, 'CHE-100.792.778', 'CHE-112.745.774', 'CHE-101.571.754', 'CHE-338.273.631', 'CHE-263.132.571', 'CHE-108.621.407', 'CHE-449.004.395', 'CHE-395.141.847', 'CHE-142.200.545', 'CHE-115.386.499', 'CHE-107.896.568', 'CHE-209.465.754', 'CHE-407.580.859', 'CHE-453.724.185']

num=1


#output filename
Output_filename='Project_Out'

dictionaries=[]

completed_list=[]

#uids=uids[:226]

print (uids)

#extracted link is joint with the url to access individual webpage

for uid in uids:
    try:
        
        dictionary={}
    
        website='https://zh.chregister.ch/cr-portal/auszug/auszug.xhtml?uid='+uid
        
        service=Service(executable_path=chromedriver_path)
    
        driver=webdriver.Chrome(service=service, options=chrome_options)
    
        driver.get(website)
    
        #driver.implicitly_wait(1)
        
        wait=WebDriverWait(driver, 5)
    
    
        Type=wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,'.col-sm-6 p:nth-child(2)')))
        Type=Type.text
        dict1={'Type':Type}
    
    
        UID=driver.find_element(By.CSS_SELECTOR,'.col-sm-4 .firmaImportantInfo')
        UID=UID.text
        dict2={'UID':UID}
    
        Business_name=driver.find_element(By.CSS_SELECTOR,'.col-md-8 .firmaImportantInfo')
        Business_name=Business_name.text
        dict3={'Business_name':Business_name}
    
        Company_addresses=driver.find_elements(By.CSS_SELECTOR,'.col-md-4 td~ td+ td')
    
        for Company_add in Company_addresses:
            if Company_add.text!="":
                Company_address=Company_add.text
        
        try:
            Company_address_list=Company_address.split('\n')

            address=Company_address_list[0]
            district=Company_address_list[1]

            dict4={'Company_address':address}
            dict5={'District':district}

        except:
            dict4={'Company_address':Company_address}
            dict5={'District':''}
    
    
        personal_data=[]
        function=[]
        signature=[]
    
        #wait=WebDriverWait(driver, 5)
    
        #elements=wait.until(EC.visibility_of_all_elements_located((By.XPATH,
    
        try:
            personal_datas=wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR,'.personen .evenRowHideAndSeek td:nth-child(4)')))
            for personal in personal_datas:
                x=personal.text
                personal_data.append(x)
        except:
            personal_data.append("")
        
        
        try:
            personal_datas=driver.find_elements(By.CSS_SELECTOR,'.personen .oddRowHideAndSeek td:nth-child(4)')
            count=1
            for personal in personal_datas:
                x=personal.text
                personal_data.insert(count,x)
                count=count+2
        except:
            pass
    
        try:
            functions_datas=wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,'.personen .evenRowHideAndSeek td:nth-child(5)')))
            for function_data in functions_datas:
                y=function_data.text
                function.append(y)
        except:
            function.append("")
                
        
        try:
            functions_datas=driver.find_elements(By.CSS_SELECTOR,'.personen .oddRowHideAndSeek td:nth-child(5)')
            count=1
            for function_data in functions_datas:
                y=function_data.text
                function.insert(count,y)
                count=count+2
        except:
            pass
    
        try:
            signature_datas=wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR,'.personen .evenRowHideAndSeek td:nth-child(6)')))
            for signature_data in signature_datas:
               z=signature_data.text
               signature.append(z)
        
        except:
            signature.append("")
            
        try:
            signature_datas=driver.find_elements(By.CSS_SELECTOR,'.personen .oddRowHideAndSeek td:nth-child(6)')
            count=1
            for signature_data in signature_datas:   
                z=signature_data.text
                signature.insert(count,z)
                count=count+2
        except:
            pass
        
        dict6={'personal_data':personal_data}
        dict7={'functions':function}
        dict8={'signature':signature}
        
        
        #appending all individual dictionaries to a larger dictionary
        dictionary.update(dict1)
        dictionary.update(dict2)
        dictionary.update(dict3)
        dictionary.update(dict4)
        dictionary.update(dict5)
        dictionary.update(dict6)
        dictionary.update(dict7)
        dictionary.update(dict8)
        
        #appending the larger dictionary to the list of dictionaries
        dictionaries.append(dictionary)
        
        #to keep a track of links which have been accessed
        completed_list.append(uid)
                    
        driver.quit()
    
    
        #Printing just to show how many completed
        print ("Number"+str(num)+" "+uid+"done--------------")
        
        num=num+1
    
    #program is run in try run block so that as soon as there is an error,
    #it breaks out and save the gathered data to excel file
    except:
        print ("error")
        break





# converting gathered data to excel file
Project_function.Data_to_Excel(dictionaries, Output_filename)

#Merge function called which will merge cells having same values
#for formatting purpose
Project_function.Merge_cells(Output_filename)

#this will export list of links accessed to excel file
#during next run, completed links will not be accessed again. 
df_completed = pd.DataFrame(completed_list, columns=['Completed_UID'])

print (completed_list)

Project_function.Convert_to_Excel(df_completed, 'Completed_UID_list')