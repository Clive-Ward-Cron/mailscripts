import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

driver = webdriver.Firefox()
driver.get("https://gateway.usps.com")

time.sleep(10)

user = driver.find_element_by_name("username")
pw = driver.find_element_by_name("password")
submit = driver.find_element_by_name("signin")
time.sleep(1)
user.clear()
time.sleep(1)
user.send_keys("clive_wardcron")
time.sleep(1)
pw.clear()
time.sleep(1)
pw.send_keys("Orion2017")

time.sleep(2)

submit.click()
