import pandas as pd
import numpy as np
import os
import time
from pynput.keyboard import Key, Controller
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
import shutil

Login_Name = "REMOVED"
Password = "REMOVED"

try:
    driver = webdriver.Chrome()
except:
    print(f"Chromedriver update required. Go to https://chromedriver.chromium.org/downloads and download version corresponding to your current version of Chrome. Then, replace chromedriver.exe in placard_automation folder with newly downloaded chromedriver.exe.")
    time.sleep(30)

driver.maximize_window()
time.sleep(2)

driver.get(f"https://jobsite.sharepoint.com/:x:/r/sites/StickersandPlacards/_layouts/15/doc2.aspx?sourcedoc=%7B5D6B9DB5-A050-45E2-8BA5-CB00AD631A82%7D&file=Sticker%20Order%20form.xlsx&action=edit&mobileredirect=true&d=w5d6b9db5a05045e28ba5cb00ad631a82&cid=4aef8337-2e93-44cc-9d05-8e9d0e3b0116")
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="i0116"]').send_keys(Login_Name)
time.sleep(1)
driver.find_element(By.ID, "idSIButton9").click()
time.sleep(1)
driver.find_element(By.ID, "i0118").send_keys(Password)
time.sleep(1)
driver.find_element(By.ID, "idSIButton9").click()

print(f'Go to File > Save As > Download a Copy.\nWindow will close 20 seconds after opening.')
time.sleep(20)
driver.close()

try:
    os.remove('Sticker Order form.xlsx')
except:
    pass

try:
    shutil.move("C:/Users/Admin/Downloads/Sticker Order form.xlsx", "C:/Users/Admin/Desktop/placard_automation/sticker_automation/")
except:
    pass

with open('completed.txt') as f:
    completed = int(f.read())

print(completed)

os.chdir("C:/Users/Admin/Desktop/placard_automation/sticker_automation/")
form = pd.read_excel(f"Sticker Order form.xlsx")
warehouses = form.iloc[:,5]
warehouses.drop(labels=range(0,completed+1), axis=0, inplace=True)
form.drop(form.columns[[0,1,2,3,4,5,6,7,8,9,11,12,34,35,37]], axis = 1, inplace=True)
form = form[form.columns[[12,0,1,2,3,4,5,6,7,8,9,10,11,13,14,15,16,17,18,19,20,21,22]]]
form.drop(labels=range(0,completed+1), axis=0, inplace=True)
form.reset_index(inplace=True)
form.drop(['index'], axis = 1, inplace=True)

orders = ''

for i in range(len(form.index)):
    order = f'Order {i+1} for {warehouses.iloc[i]}\n{form.iloc[i].loc[form.iloc[i].notna() == True]}\n\n'
    orders = orders + order

try:
    os.remove('orders.txt')
except:
    pass

with open('orders.txt', 'x') as f:
    f.write(orders)

for ind in form.index:
    os.chdir("C:/Program Files (x86)/TagPrint")
    os.system("start TagPrint.exe")
    time.sleep(15)
    keyboard = Controller()
    for i, label in enumerate(form.iloc[ind].isna()):
        if label == False:
            keyboard.press(Key.alt)
            keyboard.release(Key.alt)
            time.sleep(0.5)
            keyboard.press('f')
            keyboard.release('f')
            time.sleep(0.5)
            keyboard.press('o')
            keyboard.release('o')
            time.sleep(0.5)
            for num in str(i+1):
                keyboard.press(num)
                keyboard.release(num)
                time.sleep(0.5)
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)
            time.sleep(8)
    completed += 1

print(completed)

os.chdir('C://Users/Admin/Desktop/placard_automation/sticker_automation/')
with open('completed.txt', 'w') as f:
    f.write(str(completed))

with open('orders.txt', 'r') as f:
    if f.read() != '':
        os.system("start notepad C://Users/Admin/Desktop/placard_automation/sticker_automation/orders.txt")
