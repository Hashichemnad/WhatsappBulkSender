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

#for windows import autit
#import autoit

options = Options()
import xlrd

d = {} # dictionary to store contact details
rows=5 # no of contact details in excel

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
	message="Hi "name+"  Welcome";
  
	if number == "":
		continue
        
	try:
		url = 'https://web.whatsapp.com/send?phone=' + number + '&text=' + message
		driver.get(url)
		try:
			click_btn = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.CLASS_NAME , '_1E0Oz')))
			time.sleep(3)
			clipButton = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/div/span')
			clipButton.click()
			time.sleep(1)
			mediaButton = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[1]/button')
			mediaButton.click()
			time.sleep(1)
      
			ait.press('i', 'm', 'a','g','e')
      time.sleep(1)
			ait.press('\n')
      
#     For Windows use below code
#     image_path = os.getcwd()  + '/image.png'
# 		autoit.control_focus("Open", "Edit1")
# 		autoit.control_set_text("Open", "Edit1", image_path)
# 		autoit.control_click("Open", "Button1")

			time.sleep(1)
			whatsapp_send_button = driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div')
			whatsapp_send_button.click()
      time.sleep(2)      
            
		except (UnexpectedAlertPresentException, NoAlertPresentException) as e:
			print("alert present")
			Alert(driver).accept()
		sleep(1)
		click_btn.click()
		sleep(3)
		print('Message sent to: ' + number)
	except Exception as e:
		print('Failed to send message to ' + number + str(e))
