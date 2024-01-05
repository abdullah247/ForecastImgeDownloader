import time

from  selenium import  webdriver
from selenium.webdriver.common.by import By

from data2 import  mydata
driver =  webdriver.Chrome()


for data in mydata:
    url = data["url"]

    driver.get(url)


    while len(driver.find_elements(By.CSS_SELECTOR, "#" + data["id"])) < 1:
        a=1+3
        # print(driver.find_elements(By.CSS_SELECTOR, "#" + data["id"]), data["id"],
        #       len(driver.find_elements(By.CSS_SELECTOR, "#" + data["id"])))
    while "loading.gif" in driver.find_element(By.ID, data["id"]).get_attribute("src"):
        print(driver.find_element(By.ID, data["id"]).get_attribute("src"), "waiting for loading to go away")
        time.sleep(2)
    # if "gif" in data["name"]:
    #     print(driver.find_element(By.NAME, data["id"]).get_attribute("src").replace("ce1",
    #                                                                                 "ce" + str(data["name"][3])))
    # else:
    #     print(driver.find_element(By.ID, data["id"]).get_attribute("src"))

    name = data["name"]

    if "gif" in data["name"]:
        print(name,",",driver.find_element(By.NAME, data["id"]).get_attribute("src").replace("ce1","ce" + str(data["name"][3])))
    else:
        print(name,",",driver.find_element(By.ID, data["id"]).get_attribute("src"))


driver.close()