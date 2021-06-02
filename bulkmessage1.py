from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, NoAlertPresentException
from time import sleep
from urllib.parse import quote
from sys import platform
import os
import time
import ait

options = Options()
import xlrd

d = {} # to save the contact details
rows=5 # no of contact details in excel file
wb = xlrd.open_workbook('numbers.xls')
sh = wb.sheet_by_index(0)   
for i in range(rows):
    cell_value_class = sh.cell(i,1).value
    cell_value_id = sh.cell(i,0).value
    d[cell_value_id] = int(cell_value_class)

delay = 30


driver = webdriver.Chrome("/home/hashichemnad/Documents/whatsapp/chromedriver", options=options)
print('Once your browser opens up sign in to web whatsapp')
driver.get('https://web.whatsapp.com')
input("Press ENTER after login into Whatsapp Web and your chats are visiable	.")

for key in d.items():
	number=str(key[1]);
  name=key[0];
	message="Hi "+name+" Welcome"
  
	if number == "":
		continue
        
	try:
		url = 'https://web.whatsapp.com/send?phone=' + number + '&text=' + message

		driver.get(url)
		try:
			sleep(1)
			click_btn = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.CLASS_NAME , '_1E0Oz')))
			sleep(3)
		except (UnexpectedAlertPresentException, NoAlertPresentException) as e:
			print("alert present")
			Alert(driver).accept()
            
		sleep(1)
		click_btn.click()
		sleep(3)
		print('Message sent to: ' + number)
	except Exception as e:
		print('Failed to send message to ' + number + str(e))
