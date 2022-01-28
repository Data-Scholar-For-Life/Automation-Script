from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date, datetime, timedelta
import time
   
print('Opening Webbrowser...')
PATH = r"Pathofchromedriver.exe"
driver = webdriver.Chrome(PATH)
   
print('Accessing Webpage...')
time.sleep(10)
driver.get("https://www.webpageaddress")
   

print('Logging In...')
time.sleep(10)
driver.implicitly_wait(180)
email = driver.find_element_by_id("mat-input-0")
email.send_keys("emaillogin")
email.send_keys(Keys.RETURN)

time.sleep(10)
driver.implicitly_wait(180)
password = driver.find_element_by_id("mat-input-3")
password.send_keys("Passwordentry")
password.send_keys(Keys.RETURN)


print('Entering Site...')
time.sleep(10)
driver.implicitly_wait(180)
allcolumns = driver.find_element_by_xpath("loginbuttonxpath")
allcolumns.click()

time.sleep(10)
driver.implicitly_wait(180)
login = driver.find_element_by_id("login")
login.click()

print('Entering Search Area...')
time.sleep(10)
driver.maximize_window()
driver.implicitly_wait(180)
trysearch = driver.find_element_by_id("idofsearchpage")
trysearch.click()

time.sleep(10)
driver.implicitly_wait(180)
columnset = driver.find_element_by_id("settingcolumnsdislplayed")
columnset.click()

time.sleep(10)
driver.implicitly_wait(180)
allcolumns = driver.find_element_by_xpath("nameofcolumnset")
allcolumns.click()

print('Setting Date...')
time.sleep(10)
driver.implicitly_wait(180)
datemenu = driver.find_element_by_id("settingdatesofsearch")
datemenu.click()

time.sleep(10)
driver.implicitly_wait(18
today = driver.find_element_by_id("presettime")
today.click()

time.sleep(10)
driver.implicitly_wait(180)
datemenu = driver.find_element_by_id("closingoutdateentry")
datemenu.click()


print('Exporting...')
time.sleep(10)
driver.implicitly_wait(180)
Export = driver.find_element_by_xpath("xpathofexcelxportoption")
Export.click()

print('Export Complete...')







