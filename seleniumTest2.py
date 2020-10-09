import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

fp = webdriver.FirefoxProfile("C:/Users/digital/AppData/Roaming/Mozilla/Firefox/Profiles/o6kokyap.default")
driver = webdriver.Firefox(fp)

# driver.get("https://gateway.usps.com")
driver.get("https://anymod.com/login")

#user = driver.find_element_by_name("username")
#pw = driver.find_element_by_id("tPassword")
#submit = driver.find_element_by_name("signin")

user = driver.find_element_by_tag_name("input")
user.send_keys("clive@ward-cron.design")

pw = driver.find_element_by_xpath("//input[@type='password']")
pw.send_keys("Orion2017!")

driver.find_element_by_xpath("//button[@type='submit']").click()
