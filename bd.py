import os
import sys
import time
from base64 import main
import pyautogui
import selenium
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

npsbot = webdriver.Chrome("chromedriver.exe")
# calling a site  example

# driver.quit()
# npsbot.quit()

#  Voltas NPS Automation Bruteforce

print("Opening Links file")
npslinks = "links.txt"

npsbot.maximize_window()
with open(npslinks) as cogent:
    try:
        for num, url in enumerate(cogent):
            npsbot.get(url.strip())
            # Get All elements
            urconv = npsbot.find_element_by_id('yourcon1')
            firstv = npsbot.find_element_by_id('firstv1')
            wastechcou = npsbot.find_element_by_id('wastechcou1')
            chargingtech = npsbot.find_element_by_id('chargingtech1')
            # 3 stars
            starratting = npsbot.find_element_by_xpath('/html/body/div/div/div/div[2]/form/div/div[1]/div[7]/div[2]/div/div[1]/input')
            actions = ActionChains(npsbot)

            scalerate = npsbot.find_element_by_id('ratingpremain10')
            remarks = npsbot.find_element_by_id('remarksbreak').send_keys('good service')
            submit = npsbot.find_element_by_id('submit')

            #pyautogui.click(792, 265)
            # Do Action on elements
            urconv.click()
            firstv.click()
            wastechcou.click()
            chargingtech.click()
            #starratting.click()
            actions.click(starratting)
            actions.perform()
            scalerate.click()
            time.sleep(1)
            submit.click()
            time.sleep(2)
    except Exception as e:
        print(e)
npsbot.quit()

if __name__ == '__manin__': main()
