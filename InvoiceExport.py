import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

driver = webdriver.Edge('msedgedriver.exe')
invoicePage = "https://app.neat.com/invoices"
invoiceFile = "YearlyInvoices.xlsx"     #NOTE: Be sure to include an empty xlsx file in the project folder with this name

counter = 321   #NOTE: Replace counter number with the number of invoices you wish to pull 

invoiceData = {     #NOTE: Replace "Customer#" with short form names of your customers
    "Counter": [0],
    "Customer1": [0],
    "Customer2": [0],
    "Customer3": [0],
    "Customer4": [0],
    "Customer5": [0],
    "Customer6": [0],
    "Customer7": [0],
    "Customer8": [0],
    "Customer9": [0]
    }

invoiceDF = pd.DataFrame(invoiceData)

def login():    #Handles automatic login
    #NOTE: Replace these with your username and password
    username = "username"
    password = "password"

    userElement = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id-username"]')))
    userElement.click()
    userElement.send_keys(username)

    passElement = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id-password"]')))
    passElement.click()
    passElement.send_keys(password)

    passElement.send_keys(Keys.ENTER)

def dataUpdater():     #Pulls in data from excel that was filled during previous runs
    global invoiceDF
    global counter

    invoiceDF = pd.read_excel(invoiceFile)      #Read data from Excel files
    
    counter = min(invoiceDF["Counter"]) - 1 #Update counter to continue where the program left off

    print("---EXISTING DATA IMPORTED---")
    print(invoiceDF)

def startup():  #Pulls in previous data, loads up invoice page, and order invoices from greatest to least
    dataUpdater()   #NOTE: On the first program run disable this function call. The excel will be empty and therefore throw an error.

    #Load up the website and wait until invoice page is reached. This allows time for entering username and password.
    driver.get(invoicePage)
    login()
    while driver.current_url  != invoicePage:
        time.sleep(1)
    print("---INVOICE PAGE REACHED---")

    #Organize invoices from greatest to least
    listSort = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt-tab-panel_undefined_inv"]/div[2]/div/div/div/div/div[1]/div[4]')))
    listSort.click()
    time.sleep(1)

def loadAll():  #Scroll to the bottom of the invoices list in order to load all the invoices
    time.sleep(3)
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt-tab-panel_undefined_inv"]/div[2]/div/div/div/div/div[2]/div[1]'))).click()

    while len(driver.find_elements(By.XPATH, '//*[@id="pt-tab-panel_undefined_inv"]/div[2]/div/div/div/div/div[337]')) < 1:
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.PAGE_DOWN)

def scrapeInfo():   #Grab the invoice amount and customer name from the "View Invoice" page
    while "https://app.neat.com/invoices/preview" in driver.current_url == False:
        time.sleep(1)

    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'InvoicePreviewPage_frame__vnXeD')))

    driver.switch_to.frame(iframe)
    
    name = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/span').text
    subtotal = driver.find_element(By.XPATH, '/html/body/div/div[3]/div[1]/table/tbody/tr[1]/td[2]').text

    #All values are stored in cents to avoid floating point errors
    subtotal = subtotal.replace("$", "")
    subtotal = subtotal.replace(".", "")
    print(name)
    print(subtotal)

    moneyTracker(name, subtotal)
    
    driver.back()

    while driver.current_url  != invoicePage:
        time.sleep(1)

def moneyTracker(name, subtotal):   #Keeps track of total invoice amount for each customer
    global invoiceDF
    
    #NOTE: Fill "Customer# Full Name" with the full name of your customers as they appear on Neat. Ensure that the order matches the one's previously set.
    match name:
        case "Customer1 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": subtotal,
                "Customer2": 0,
                "Customer3": 0,
                "Customer4": 0,
                "Customer5": 0,
                "Customer6": 0,
                "Customer7": 0,
                "Customer8": 0,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer2 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": subtotal,
                "Customer3": 0,
                "Customer4": 0,
                "Customer5": 0,
                "Customer6": 0,
                "Customer7": 0,
                "Customer8": 0,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer3 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": 0,
                "Customer3": subtotal,
                "Customer4": 0,
                "Customer5": 0,
                "Customer6": 0,
                "Customer7": 0,
                "Customer8": 0,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer4 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": 0,
                "Customer3": 0,
                "Customer4": subtotal,
                "Customer5": 0,
                "Customer6": 0,
                "Customer7": 0,
                "Customer8": 0,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer5 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": 0,
                "Customer3": 0,
                "Customer4": 0,
                "Customer5": subtotal,
                "Customer6": 0,
                "Customer7": 0,
                "Customer8": 0,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer6 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": 0,
                "Customer3": 0,
                "Customer4": 0,
                "Customer5": 0,
                "Customer6": subtotal,
                "Customer7": 0,
                "Customer8": 0,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer7 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": 0,
                "Customer3": 0,
                "Customer4": 0,
                "Customer5": 0,
                "Customer6": 0,
                "Customer7": subtotal,
                "Customer8": 0,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer8 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": 0,
                "Customer3": 0,
                "Customer4": 0,
                "Customer5": 0,
                "Customer6": 0,
                "Customer7": 0,
                "Customer8": subtotal,
                "Customer9": 0
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)
        case "Customer9 Full Name":
            dataAppend = {
                "Counter": counter,
                "Customer1": 0,
                "Customer2": 0,
                "Customer3": 0,
                "Customer4": 0,
                "Customer5": 0,
                "Customer6": 0,
                "Customer7": 0,
                "Customer8": 0,
                "Customer9": subtotal
            }
            dfAppend = pd.DataFrame(dataAppend, index=[0])
            invoiceDF = pd.concat([invoiceDF, dfAppend],axis=0)

startup()

#Iterate through all the button elements and view each one
while counter >= 0:
    loadAll()
    time.sleep(0.25)
    moreElements = driver.find_elements(By.CLASS_NAME, 'actions-dropdown')

    while len(moreElements) < counter + 10:
        time.sleep(1)
        moreElements = driver.find_elements(By.CLASS_NAME, 'actions-dropdown')
    
    print(f"Counter: {counter}")

    #Need to use page up for scrolling the last few buttons into view
    if counter > 5:
        driver.execute_script("arguments[0].scrollIntoView();", moreElements[counter - 3])
    
    else:
        driver.execute_script("arguments[0].scrollIntoView();", moreElements[5])
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.PAGE_UP)
    
    moreElements[counter] = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(moreElements[counter]))
    moreElements[counter].click()

    viewElement = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'View')]")))
    viewElement.click()

    scrapeInfo()
    invoiceDF.to_excel(invoiceFile)

    counter -= 1
    print("#####################")

print("-------DONE-------")
time.sleep(5)
print("CLOSING DRIVER")
driver.close()