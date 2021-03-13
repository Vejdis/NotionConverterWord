from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time

import tkinter as tk
from tkinter import simpledialog

#Selenium setup
PATH = "E:\Libraries\Programmering\Selenium\chromedriver.exe"
driver = webdriver.Chrome(PATH)

#Create simple dialogbox for what file we want to export
ROOT = tk.Tk()
ROOT.withdraw()
fileName = simpledialog.askstring(title = "Notion File", prompt = "What file to convert")

#Open notion page
driver.get("NOTION START PAGE HERE")

#Wait for page to load, sometimes unneccesary but annoying when selenium can't find element
time.sleep(2)

#Enter email
email = driver.find_element_by_id("notion-email-input-1")
email.send_keys("YOUR EMAIL HERE")
email.send_keys(Keys.RETURN)

#Wait for page to load, sometimes unneccesary but annoying when selenium can't find element
time.sleep(2)

#Enter password
pw = driver.find_element_by_id("notion-password-input-2")
pw.send_keys("PASSWORD HERE")
pw.send_keys(Keys.RETURN)

#Wait for page to load, sometimes unneccesary but annoying when selenium can't find element
time.sleep(5)

#Open the search by shortcut control+p
openSearch = ActionChains(driver)
openSearch.key_down(Keys.CONTROL).send_keys("p").key_up(Keys.CONTROL).perform()

#sleep
time.sleep(5)

#enter search from box we started with
writeSearch = ActionChains(driver)
writeSearch.send_keys(fileName)
writeSearch.perform()

#sleep
time.sleep(5)

openFile = ActionChains(driver)
openFile.send_keys(Keys.RETURN)
openFile.perform()

#sleep
time.sleep(5)

#Open menu
menuDots = driver.find_element_by_class_name("dots").click()

#sleep
time.sleep(5)

#Click export
openExport = driver.find_element_by_xpath("//*[contains(text(), 'Export')]").click()

#sleep
time.sleep(5)

#Download export
downloadExport = driver.find_element_by_xpath("//*[contains(text(), 'Export')]").click()

#sleep
time.sleep(5)

#Run powershell script. 
# Script renames the first file in downloads folder (which should be the file just exported from notion if sorted on date modified.)
# Script moves folder to specified location
# Extracts file from zip
# Converts the markdown file to word and names it File.docx
# Deletes the .zip and the markdown file.
import os
psPath = "POWERSHELLSCRIPT PATH HERE"
os.startfile(psPath)

#close browser
driver.quit()