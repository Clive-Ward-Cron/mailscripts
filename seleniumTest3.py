import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

#fp = webdriver.FirefoxProfile("C:/Users/digital/AppData/Roaming/Mozilla/Firefox/Profiles/o6kokyap.default")
#driver = webdriver.Firefox(fp)

#driver = webdriver.Chrome()

driver = webdriver.Firefox()

driver.get("https://gateway.usps.com")
# driver.get("https://reg.usps.com/entreg/LoginAction_input?app=EDDM&appURL=https://eddm.usps.com/eddm/")


user = driver.find_element_by_name("username")
pw = driver.find_element_by_name("password")
#submit = driver.find_element_by_id("btn-submit")
#form = driver.find_element_by_id("loginForm")

user.clear()
user.send_keys("clive_wardcron")

pw.clear()
pw.send_keys("Orion2017")

time.sleep(2)
#form.submit()
#submit.click()

