# -*- coding: utf-8 -*-
"""
Created on Tue May 17 14:47:36 2022

@author: gou
"""

import pandas as pd
from selenium import webdriver
from time import sleep
import numpy as np
from selenium.common.exceptions import NoSuchElementException
#%%


df = pd.read_excel("M:\List Matching\Axles.xlsx")
df = df.iloc[:,:6]
df[['Feature_Names','Feature1', 'Feature2', 'Feature3', 'Feature4',
    'Feature5', 'Feature6', 'Feature7', 'Feature8', 'Feature9']] = np.NAN

#%%
df = pd.read_excel("M:\List Matching\Axle_Row166.xlsx", index_col=(0))

#%%

driver = webdriver.Chrome(r'C:\Users\gou\Downloads\chromedriver_win32\chromedriver.exe')
driver.implicitly_wait(10)

# For Loop
#for i in range(len(df)):
for i in range(525,len(df)):
    driver.get(df['URL'].iloc[i])
    sleep(1)
    # Seller's ID
    seller = np.NAN
    try:
        seller = driver.find_element_by_xpath('//span[@class="ux-textspans ux-textspans--PSEUDOLINK ux-textspans--BOLD"]').text
    except:
        try:
            seller = driver.find_element_by_xpath('//div[@class="pt20"]').text
            seller = seller[seller.find(" by ")+4:]
        except NoSuchElementException:
            pass
    df['Seller ID'].iloc[i] = seller
    sleep(1)
    
    # MMY
    try:
        names = []
        headers = driver.find_elements_by_xpath('//table/tbody/tr/th')
        for he in headers:
            names.append(he.text)
            df['Feature_Names'].iloc[i] = ','.join(map(str,names))
    
        for j in range(len(names)):
            df.iloc[i,j+8] = driver.find_element_by_xpath('//table/tbody/tr[3]/td[{}]'.format(j+1)).text
        sleep(1)
    except NoSuchElementException:
        pass
    
    print(i, " is completed")
driver.quit()
#%%
df.to_excel("M:\List Matching\Axle_Completed.xlsx")



































