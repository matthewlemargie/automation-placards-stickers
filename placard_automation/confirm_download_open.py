from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
import os
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import shutil
from pynput.keyboard import Key, Controller

Login_Name = "REMOVED"
Password = "REMOVED"

placards = pd.read_csv('placards.csv')
placards = placards[['name', 'warehouse', 'id', 'staged']]

staged = placards.loc[placards['staged'] == 1]


# confirm completed placards
try:
    driver = webdriver.Chrome()
except:
    print(f"Chromedriver update required. Go to https://chromedriver.chromium.org/downloads and download version corresponding to your current version of Chrome. Then, replace chromedriver.exe in placard_automation folder with newly downloaded chromedriver.exe.")
driver.maximize_window()
driver.get("https://companysite.com/Login")
driver.find_element(By.ID, "Login_Name").send_keys(Login_Name)
driver.find_element(By.ID, "Password").send_keys(Password)
driver.find_element(By.ID, "loginButton").click()
time.sleep(2)

bad_placards = pd.read_csv('bad_placards.csv', header=None)
bad_placards = bad_placards[0].array

for id in staged.id:
    try:
        # confirm placard on website
        driver.get(f"https://companysite.com/CADPlacard/Edit/{id}")
        time.sleep(1)
        driver.find_element(By.ID, "Job_PlacardPrinted").click()
        time.sleep(0.2)
        driver.find_element(By.ID, "Job_PlacardSent").click()
        time.sleep(0.2)
        # double click save button
        save_button = driver.find_element(By.XPATH, '//*[@id="SaveBtn"]')
        actionChains = ActionChains(driver)
        actionChains.double_click(save_button).perform()
        time.sleep(0.2)
    except:
        print(f"{id} is bad")
        f = open("bad_placards.csv", 'a')
        f.write(f'\n{id}')
        f.close()
    
    
    # move from queue to complete folder
    files = [pdf for pdf in os.listdir(f'C://Users//Admin//Desktop//Placard Q//Queue')]
    for file in files:
        if placards.name.iloc[placards[placards['id'] == id].index[0]] in file:
            shutil.move(f'C://Users//Admin//Desktop//Placard Q//Queue//{file}', f'C://Users//Admin//Desktop//Placard Q//Complete//{file}')
                   
# remove from placards.csv         
for id in staged.id:
    placards = placards.drop(placards[placards.id == id].index, axis=0)

# update placards.csv                   
placards.reset_index()[['name', 'warehouse', 'id', 'staged']].to_csv('placards.csv')


# download new placards

# go to placard website
driver.get("https://companysite.com/CADPlacard")
time.sleep(1)

# go to new placards
driver.find_element(By.XPATH, '//*[@id="#renderbody"]/div[2]/div[1]/div[1]/div/div/div/div/a[3]').click()
time.sleep(1)

#lengthen placard list
sel = Select(driver.find_element(By.XPATH, '//*[@id="loiTable_length"]/label/select'))
sel.select_by_value('25')

# create soup
soup = BeautifulSoup(driver.page_source, 'html.parser')
body_rows = soup.find('table', id = 'loiTable').find("tbody").findAll('tr')


placards = pd.read_csv('placards.csv')

init_placards = pd.read_csv('placards.csv')
init_placards = init_placards['id'].array

init_files = [file for file in os.listdir(f'C://Users//Admin//Desktop//Placard Q//Queue')]
init_files.remove('email teams etc')

#placards = pd.DataFrame(columns=['name', 'warehouse', 'id', 'staged'])

#download first x pdfs from placard list

#ask how many to download
download = None
while not isinstance(download, int):
    try:
        download = int(input(f"Enter how many placards to be downloaded: "))
    except:
        print(f'Please enter an integer.')

#add to download total to account for skipped placards
for row in body_rows[:download]:
    placard_id = row['id']

    if placard_id in bad_placards:
        download += 1

files = [file for file in os.listdir(f'C://Users//Admin//Desktop//Placard Q//Queue')]
files.remove('email teams etc')

for file in files:
    download += 1

#download from cadplacard list
try:
    for row in body_rows[:download]:


        name = row.findAll('td')[2].text.replace('  ', '').replace('\n', '')
        warehouse = row.findAll('td')[4].text.replace('  ', '').replace('\n', '')
        placard_id = row['id']

        # skip unremovable placards
        if placard_id in bad_placards:
            continue

        # skip already downloaded placards
        if placard_id in init_placards:
            continue

        pdf_dir = row.findAll('td')[9].find('a')['href']

        placard = pd.DataFrame({'name': [name], 'warehouse': [warehouse], 'id': [placard_id], 'staged': [0]})
        placards = pd.concat([placard, placards])
        
        url = f'https://companysite.com{pdf_dir}'
        pdf = requests.get(url)
        
        direc = f'C://Users//Admin//Desktop//Placard Q//Queue'
        direc_list = [file for file in os.listdir(direc)]
        file_name = f'{name} {warehouse}.pdf'

        # prevents program from rewriting over placards with the same name
        # this ensures they are saved with different names

        i = 1
        while file_name in direc_list:
            file_name = f'{name}{i} {warehouse}.pdf'
            i += 1

        file_path = f'{direc}//{file_name}'

        with open(file_path, 'wb') as f:
            f.write(pdf.content)
            
        placards.drop_duplicates(['name'], inplace=True)

    placards[['name', 'warehouse', 'id', 'staged']].to_csv('placards.csv')
except:
    print(f'No placards were downloaded') 

driver.close()

# open downloaded pdfs in queue folder
files = [file for file in os.listdir(f'C://Users//Admin//Desktop//Placard Q//Queue')]
files.remove('email teams etc')
for file in init_files:
    files.remove(file)
keyboard = Controller()

# open and select all elements in each file
os.system(f'start Acrobat.exe')
time.sleep(1)
for file in files:
    # open pdf
    os.system(f'start Acrobat.exe C://Users//Admin//Desktop//Placard Q//Queue//{file}')
    time.sleep(1)
    
    # enter editing mode
    keyboard.press(Key.alt)
    keyboard.release(Key.alt)
    time.sleep(0.3)
    keyboard.press('v')
    keyboard.release('v')
    time.sleep(0.3)
    keyboard.press('t')
    keyboard.release('t')
    time.sleep(0.3)
    keyboard.press(Key.down)
    keyboard.release(Key.down)
    time.sleep(0.3)
    keyboard.press(Key.down)
    keyboard.release(Key.down)
    time.sleep(0.3)
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(0.3)
    keyboard.press('o')
    keyboard.release('o')
    
    time.sleep(1)
    
    # select all elements in pdf
    keyboard.press(Key.shift)
    keyboard.press(Key.f10)
    time.sleep(0.3)
    keyboard.release(Key.f10)
    keyboard.release(Key.shift)
    time.sleep(0.3)
    keyboard.press('l')
    keyboard.release('l')
    time.sleep(0.3)
    
    # zoom to page level
    keyboard.press(Key.ctrl)
    keyboard.press('0')
    time.sleep(0.3)
    keyboard.release(Key.ctrl)
    keyboard.release('0')
    
    time.sleep(1)
