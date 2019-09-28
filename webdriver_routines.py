from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select

from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options

#this file will contain the general routines used in using a webdriver
#all the functions will take a webdriver object and lalala.



def print_element_texts(elements):
    text_list = []
    i = 0
    for element in elements:
        print(str(i) + ': ' + element.text)
        i = i + 1
        text_list.append(element.text)

def chromedriver_initialize(driver_file_path, headless = True, images = True):
    #this is the location of the chromedriver file
    #https://sites.google.com/a/chromium.org/chromedriver/downloads
    #driver_file_path= "C:\\Users\\Leslie\\Documents\\Programming\\chromedriver.exe"
    options = webdriver.ChromeOptions()

    #TURN THIS OPTION ON TO ENABLE HEADLESS
    if headless == True:
        options.add_argument('headless')

    options.add_argument('window-size=1200x600') # optional
    browser = webdriver.Chrome(executable_path=driver_file_path, options=options)
        
    return browser
