# -*- coding: utf-8 -*-
"""
Created on Thu Jan 12 10:58:51 2023

@author: shrihari
"""
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as  pd
from datetime import datetime
import time
import shutil
import credfile
import os
import sys

import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

import paths
from error_log import log_error


STANDARD_HEADERS = ['Payee/Vendor', 'Client Claim Number', 'Invoice Number', 'Assignment Type', 'Adjuster Cost Category', 'Grand Total']

LOCAL_DOWNLOADS = paths.local_download_path
LOCAL_OUTPUT = paths.local_completed_path
LOCAL_ERROR = paths.local_error_path

def call_process(fName, credentials):
    startTime = time.time()

    try:
        usr = credentials['username']
        pwd = credentials['password']
        base_name, ext = os.path.splitext(fName)
        current_time = datetime.now().strftime('%m_%d_%Y')
        filename = f"{base_name}_{current_time}{ext}"
        filename1 = fName
        filename1_flag = 0
        exception_flag = 0

        source_filepath = os.path.join(LOCAL_DOWNLOADS, fName)
        destination_filepath = os.path.join(LOCAL_DOWNLOADS, filename)
        shutil.move(source_filepath, destination_filepath)
        data = pd.read_excel(destination_filepath)
        missing_headers = set(STANDARD_HEADERS) - set(list(data.columns))
        print(f"\nMissing headers: {missing_headers}")
        if missing_headers == {}:
            print("Header Validation - success")
            
            data['EOX Comments'] = ''
            
            def click_using_text(textValue):
                # Replace 'Click me' with the text of the element you want to click
                element_text = textValue
                
                # Find and click the element based on its text
                found = False
                for element in driver.find_elements(by=By.XPATH,value="//*[text() = '%s']" % element_text):
                    try:
                        element.click()
                        found = True
                        break  # Stop searching once we've found and clicked the element
                    except Exception as e:
                        pass  # Ignore exceptions if the click fails for this element
        
            def wait_for_ele_presence(ele,timelimit=30):
                WebDriverWait(driver,timelimit).until(EC.presence_of_element_located((By.XPATH,ele)))
            # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            
            def search_claim_number_in_webpage(cNum):
                searchButton = '//*[@id="header"]/div/div[1]/div[1]/div'
                wait_for_ele_presence(searchButton,100)
                driver.find_element(by=By.XPATH,value=searchButton).click()
                
                inputField = '//*[@id="header"]/div/div[1]/div[1]/input'
                wait_for_ele_presence(inputField)
                driver.find_element(by=By.XPATH,value=inputField).send_keys(cNum,Keys.ENTER)
                time.sleep(2)
            
            def check_for_webpage_text(checkTextWebpage,waitTime=10):
                webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                #print(webpageText)
                countwebpageText = 0
                while checkTextWebpage not in webpageText and countwebpageText <= waitTime:
                    time.sleep(1)
                    webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                    countwebpageText += 1
            
            def Change_Claim_Status(current,toChange):
                claimDetails = '//*[@id="scaffold-wrapper"]/div/div[1]/div/div[4]'
                wait_for_ele_presence(claimDetails,60)
                claimDetailsCheckText = driver.find_element(by=By.XPATH,value=claimDetails).text
                countclaimDetailsCheckText = 0
                while claimDetailsCheckText != 'Claim Details' and countclaimDetailsCheckText <= 30 :
                    claimDetailsCheckText = driver.find_element(by=By.XPATH,value=claimDetails).text
                    time.sleep(0.5)
                    countclaimDetailsCheckText += 1
                    #print(countclaimDetailsCheckText,claimDetailsCheckText)
                driver.find_element(by=By.XPATH,value=claimDetails).click()
                
                closedCheckText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]/div[1]/div[3]/div/div/label/div/div/div[1]/div[1]/div').text
                countclosedCheckText = 0
                while closedCheckText.upper() != current.upper() and countclosedCheckText <=25:
                    closedCheckText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]/div[1]/div[3]/div/div/label/div/div/div[1]/div[1]/div').text
                    time.sleep(1)
                    countclosedCheckText += 1
                    #print(countclosedCheckText,closedCheckText)
                time.sleep(1)
                driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]/div[1]/div[3]/div/div/label/div/div').click()
                driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[1]/div[1]/div[3]/div/div/label/div/div/div[1]/div[2]/div/input').send_keys(toChange,Keys.ENTER)
                
                
            # Enabling Driver options
            options = Options()
            options.add_experimental_option("excludeSwitches" , ["enable-automation"])
            # options.add_argument('--headless')
            options.add_argument("--window-size=1920,1080")
            chrome_path = r"chromedriver.exe" #path from 'which chromedriver'#path from 'which chromedriver'
            # chrome_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(options=options)
            driver.maximize_window()

            # driver.get('https://mail.google.com')
            # cookies = get_cookies_from_website(driver, "https://mail.google.com")
            # time.sleep(25)
            # driver = add_cookies_to_website(driver, 'https://test.snapsheetvice.com', cookies)

            try:
                driver.get(credfile.url)
                print('added')
                
            except:
                time.sleep(60)
                print("Issue While Opening the Browser")
                driver.get(credfile.url)
            #driver.maximize_window()
            driver.implicitly_wait(15)
            
            # =============================================================================
            # **************************   login ****************************************
            # =============================================================================
            driver.find_element(by=By.XPATH,value='//*[@id="app"]/div/form[2]/label/input').send_keys(usr,Keys.ENTER)
            time.sleep(5)
            driver.find_element(by=By.XPATH,value='//*[@id="app"]/div/form[2]/label[2]/input').send_keys(pwd)
            time.sleep(5)
            driver.find_element(by=By.XPATH,value='//*[@id="app"]/div/form[2]/label[2]/input').send_keys(Keys.ENTER)
            
            # element_location = element.location
            # element_size = element.size
            
            # # Calculate the center position of the element
            # center_x = element_location['x'] + element_size['width'] / 2
            # center_y = element_location['y'] + element_size['height'] / 2
            
            # # Move the cursor slowly to the element using pyautogui
            # move_duration = 2  # Duration in seconds
            # pyautogui.moveTo(center_x, center_y, duration=move_duration)
            
            for pos, (index, row) in enumerate(data.iterrows()):
                checkPointPayment = 1
                try:
                    # =============================================================================
                    # ******************  search and enter the claim value ***********************
                    # =============================================================================
                    clnumberfromexcel = row['Client Claim Number']
                    clnumberfromexcel = str(clnumberfromexcel).replace(',','')
                    clnumberfromexcel = clnumberfromexcel.strip()
                    clnumberfromexcel = clnumberfromexcel.upper()
                    invoiceNumber = row['Invoice Number']
                    #print(clnumberfromexcel)
                    
                    search_claim_number_in_webpage(clnumberfromexcel)
                    
                    #print('search time',time.time() - startTime)
                    
                    claimsListCount = driver.find_elements(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div/div[3]/ul/div/div[3]/div/div[2]/div')
                    
                    for clLen in range(len(claimsListCount)):
                        try:
                            claimNumWeb = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div/div[3]/ul/div/div[3]/div/div[2]/div['+str(clLen+1)+']/div[1]/div[2]/a/span').text
                        except:
                            time.sleep(3)
                            claimNumWeb = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div/div[3]/ul/div/div[3]/div/div[2]/div['+str(clLen+1)+']/div[1]/div[2]/a/span').text
                        
                        if claimNumWeb == clnumberfromexcel:
                            print(claimNumWeb)
                            claimStatus = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div/div[3]/ul/div/div[3]/div/div[2]/div['+str(clLen+1)+']/div[2]/div[2]/div').text
                            time.sleep(1)
                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div/div[3]/ul/div/div[3]/div/div[2]/div['+str(clLen+1)+']/div[1]/div[2]/a/span').click()
                            break
                        
                    # =============================================================================
                    # --------------------- Claims Details Handling -------------------------------
                    # =============================================================================
                    if claimStatus.upper() != 'CANCELLED':
                        if claimStatus == 'Closed':
                            webpageTextWaitCounter = 0
                            webpageTextWait = driver.find_element(by=By.XPATH,value='/html/body').text
                            time.sleep(3)
                            if claimNumWeb not in webpageTextWait and webpageTextWaitCounter <=20:
                                time.sleep(1)
                                webpageTextWait = driver.find_element(by=By.XPATH,value='/html/body').text
                                webpageTextWaitCounter+=1
                                
                            Change_Claim_Status('Closed','Open')
                            
                            webpageTextWaitCounter1 = 0
                            webpageTextWait = driver.find_element(by=By.XPATH,value='/html/body').text
                            time.sleep(3)
                            if 'OPEN' not in webpageTextWait and webpageTextWaitCounter1 <=20:
                                time.sleep(1)
                                webpageTextWait = driver.find_element(by=By.XPATH,value='/html/body').text
                                webpageTextWaitCounter1+=1
                            
                            yesCheckText = driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').text
                            countyesCheckText = 0
                            while yesCheckText != 'Yes' and countyesCheckText <= 20:
                                yesCheckText = driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').text
                                time.sleep(1)
                                countyesCheckText += 1
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').click()
                            # print('Claim Reopend')
                            
                        
                        # =============================================================================
                        #  ------------------------ Claim Summary Process -----------------------------
                        # =============================================================================
                        # print('Starting CLaim Summary')
                        claimSummaryText = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[1]/div/div[2]').text
                        countclaimSummaryText = 0
                        while claimSummaryText != 'Claim Summary' and countclaimSummaryText < 10:
                            claimSummaryText = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[1]/div/div[2]').text
                            time.sleep(1)
                            countclaimSummaryText += 1
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        countwebpageText = 0
                        while ('CLAIM: '+clnumberfromexcel not in webpageText) and countwebpageText <= 20:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText += 1
                                    
                        
                        driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[1]/div/div[2]').click()
                        ClaimsSummaryListTable = '//*[@id="claim-page-wrapper"]/div/div/div[3]/div[2]/table/tbody'
                        ClaimsSummaryListTableText = driver.find_element(by=By.XPATH,value=ClaimsSummaryListTable).text
                        # print('Clicked CLaim Summary')
                        
                        if 'Dwelling' in ClaimsSummaryListTableText:
                            firstexposureText = 'Dwelling'
                            cliamSummaryListRows = driver.find_elements(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div[3]/div[2]/table/tbody/tr')
                            for cRow in range(len(cliamSummaryListRows)):
                                dwellingCheckText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div[3]/div[2]/table/tbody/tr['+str(cRow+1)+']/td[2]/button').text
                                if dwellingCheckText == 'Dwelling':
                                    time.sleep(1)
                                    # Scroll the element into view
                                    element = driver.find_element(By.XPATH, '//*[@id="claim-page-wrapper"]/div/div/div[3]/div[2]/table/tbody/tr['+str(cRow+1)+']/td[2]/button')
                                    driver.execute_script("arguments[0].scrollIntoView(true);", element)
                                    driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div[3]/div[2]/table/tbody/tr['+str(cRow+1)+']/td[2]/button').click()
                                    vendorButtonXpath = '//*[@id="claim-page-wrapper"]/div/div[1]/div/div[2]/a[3]/button'
                                    break
                        else:
                            firstexposureText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div[3]/div[2]/table/tbody/tr[1]/td[2]/button').text
                            driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div[3]/div[2]/table/tbody/tr[1]/td[2]/button').click()
                            
                            if firstexposureText == 'Personal Property':
                                vendorButtonXpath = '//*[@id="claim-page-wrapper"]/div/div[1]/div/div[2]/a[3]/button'
                            elif firstexposureText == 'Other Structures':
                                vendorButtonXpath = '//*[@id="claim-page-wrapper"]/div/div[1]/div/div[2]/a[2]/button'
                            else:
                                countOfVEndorCheckNames = driver.find_elements(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[2]/a')
                                
                                # =============================================================================
                                #         #fill in for looping to get vendor button for other cases
                                # =============================================================================
                        
                            
                
                        # =============================================================================
                        #  ---------------------------> Adding Vendors  <------------------------------
                        # =============================================================================
                        vendorNameExcel = row['Payee/Vendor']
                        driver.find_element(by=By.XPATH,value=vendorButtonXpath).click()
                        vendorTableHeader = '//*[@id="claim-page-wrapper"]/div/div[2]/div[2]/div[1]'
                        wait_for_ele_presence(vendorTableHeader)
                        
                                    # *********** check if exposure is open/closed ******************
                        
                        exposureOpenCloseTextXpath = '//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]/div[1]/div[1]/div[3]/div/label/div/div'
                        wait_for_ele_presence(exposureOpenCloseTextXpath,35)
                        exposureOpenCloseText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]/div[1]/div[1]/div[3]/div/label/div/div').text
                        
                        if exposureOpenCloseText == 'CLOSED':
                            driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]/div[1]/div[1]/div[3]/div/label/div/div').click()
                            driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[1]/div[1]/div[1]/div[3]/div/label/div/div[1]/div[1]/div[2]/div/input').send_keys('Open',Keys.ENTER)
                            
                            check_for_webpage_text('Reopen Exposure')
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').click()
                            countexposureOpenCloseText = 0
                            while exposureOpenCloseText != 'OPEN' and countexposureOpenCloseText <= 120:
                                time.sleep(0.5)
                                exposureOpenCloseText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]/div[1]/div[1]/div[3]/div/label/div/div').text
                                countexposureOpenCloseText += 1
                        try:
                            try:#get the vendor table using relative xpath when it failes use the class name
                                table = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[2]/div[2]/div[1]')
                            except:
                                table = driver.find_element(by=By.XPATH,value='//div[@class="_5UQt0Iza1XW9Hs74zdxLd nCHYD7xlq6UE1q-7xb1G9 _2buL7UDmuAvUMZo4rGgNdA"]')
                            # Find all the rows in the table
                            # print("table:\t\t",table)
                            rows = table.find_elements(by=By.XPATH,value='//a[@class="_38ZQikJDvcmVI21d2WY6T"]')
                            # print("rows: \n\n",rows)
                            # Extract the second column of data from each row
                            column_data = []
                            for row in rows:
                                # Find the second column using XPath
                                column = row.find_element(by=By.XPATH,value='//div[@class="_2DfN-63Z9Qg88hTLEmZtST _10O5tZHcEzLO4SQWu3_5fK"]')
                                # Get the text from the column and append it to the list
                                column_data.append(column.text)
                            
                            # Get the column data of vendors
                            vendorTable =  [column_data[i:i+2] for i in range(0, len(column_data), 2)]
                            vendorTableDf = pd.DataFrame(vendorTable,columns=['Service','Vendor'])
                        except:
                            vendorTableDf = pd.DataFrame(columns=['Service', 'Vendor'])
                        
                        if vendorNameExcel not in vendorTableDf['Vendor'].values:
                            driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[2]/div[1]/button').click()
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div/div/div[1]').click()
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/div/div[1]/button[1]').click()
                            #wait for table to load
                            WebDriverWait(driver,90).until(EC.invisibility_of_element_located((By.CSS_SELECTOR,'body > div.ss-dialog-container > div.ss-modal.-iGS5v22bgSXyAcLo9mXB > div.ss-modal__content._1zuMdpB8CL0L8xzY0EnZlC > div > div > div._5UQt0Iza1XW9Hs74zdxLd._3b99HJHJ-LGaKXhJf7vHLN > div._3BCBfM2jwpuAWiCGLF9faf > div')))
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div/div[1]/div[1]/div/input').send_keys(vendorNameExcel,Keys.ENTER)
                            vendorCheckFlag = 0
                            wait_for_ele_presence('/html/body/div[3]/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/div/div[1]',30)
                            time.sleep(2.5)
                            vendorNameWeb = driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/div/div[1]').text
                            countvendorNameWeb = 0
                            
                            while vendorNameWeb != vendorNameExcel and countvendorNameWeb < 7:
                                time.sleep(1)
                                vendorNameWeb = driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/div/div[1]').text
                                #print(vendorNameWeb,vendorNameExcel,'inside the loop')
                                countvendorNameWeb += 1
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/div/div[1]').click()
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/div/div[1]/button[1]').click()
                            
                            #Fill the vendor details
                            
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div[3]/div/div[3]/div/div/label[1]').click()
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div/div/label/div/div/div[1]/div[2]').click()
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div/div/label/div/div/div[1]/div[2]/div/input').send_keys('Payable Vendor',Keys.ENTER,Keys.TAB,Keys.TAB)
                            time.sleep(1)
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[2]/div/div[5]/div/div[2]/label[1]/div[1]').click()
                            time.sleep(2)
                            #save button
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/div/div[1]/button[1]').click()
                        
                        vendorCheckFlag = 1
                        # =============================================================================
                        # --------------------------- > Financials < ---------------------------------
                        # =============================================================================
                        # check_for_webpage_text('FINANCIAL SUMMARY')
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        countwebpageText = 0
                        while 'FINANCIAL SUMMARY' not in webpageText and countwebpageText <= 15:
                            try:
                                driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[1]/div/div[2]').click()
                                driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[1]/div/div[8]').click()
                            except:
                                pass
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText += 1
                        # driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[1]/div/div[8]').click()
                        
                        reservesAddButtonXpath = '//*[@id="claim-page-wrapper"]/div/div/div[2]/div/table/thead/tr/td[2]/div/button'
                        wait_for_ele_presence(reservesAddButtonXpath)
                        driver.find_element(by=By.XPATH,value=reservesAddButtonXpath).click()
                        
                        # =============================================================================
                        #  ------------------------> Adding Reserves <---------------------------------
                        # =============================================================================
                        
                        check_for_webpage_text('Add Reserve')
                        addReserveCategory = row['Adjuster Cost Category']
                        addReserverNewValue = row['Grand Total']
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        countwebpageText = 0
                        while 'Add Reserve' not in webpageText and countwebpageText <= 15:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText += 1
                        wait_for_ele_presence('//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[1]/button',20)
                        driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[1]/button').click()
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        


                        countwebpageText1 = 0
                        while firstexposureText not in webpageText and countwebpageText1 <= 15:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText1 += 1
                        click_using_text(firstexposureText)
                        countwebpageText2 = 0
                        while 'Cost Type' not in webpageText and countwebpageText2 <= 15:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText2 += 1
                        
                        element_to_hover_over = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[1]/button')
                        actions = ActionChains(driver)
                        actions.move_to_element(element_to_hover_over)
                        actions.perform()
                        
                        
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/div/label/div/div/div[1]/div[2]').click()
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/div/label/div/div/div[1]/div[2]/div/input').send_keys(
                            'Adjusting',Keys.ENTER,Keys.TAB,
                            str(addReserveCategory),Keys.ENTER,Keys.TAB,
                            str(addReserverNewValue),Keys.ENTER,Keys.TAB,
                            Keys.TAB,Keys.TAB,Keys.TAB,Keys.TAB,Keys.ENTER)
                        #print('Added')
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        countwebpageText3 = 0
                        
                        while 'RESERVE DETAILS' not in webpageText and countwebpageText3 <= 25:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText3 += 1
                            #print('waiting for reserves details')
                            
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        countwebpageText4 = 0
                        while 'RESERVE DETAILS' in webpageText and countwebpageText4 <= 20:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText4 += 1
                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[1]/div/div[8]').click()                    
                            #print('waiting for Payments button')
                            
                        # =============================================================================
                        # ------------------------> Payments <----------------------------------------
                        # =============================================================================

                        
                        driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div[2]/div/table/thead/tr/td[4]/div/button').click()
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        countwebpageText = 0
                        while 'Primary Payee Type' not in webpageText and countwebpageText <= 15:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText += 1
                        
                        
                        #Primary Payee Type
                        driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[2]/div/div[1]/label/div/div/div[1]/div[2]').click()
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[2]/div/div[1]/label/div/div/div[1]/div[2]/div/input').send_keys('Vendor',Keys.ENTER,Keys.TAB,vendorNameExcel,Keys.ENTER,Keys.TAB,'Self Select',Keys.ENTER)
                        wait_for_ele_presence('//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[4]/div[3]/div/div/label[1]/div',20)
                        driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[4]/div[3]/div/div/label[1]/div').click()
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[4]/div[2]/label/div/div/input').send_keys(str(invoiceNumber),Keys.TAB,Keys.TAB,Keys.TAB,Keys.TAB,Keys.TAB,Keys.TAB)
                        driver.find_element(by=By.XPATH, value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[13]/div[3]/div[1]/div/div/label[2]/div').click()               
                        driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[13]/div[4]/div[1]/div/div/label[1]/div').click()
                        
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div/div[2]/div/div/div/div[4]/button[2]').click()
                        
                        
                        countwebpageText5 = 0
                        while 'Notification Only' in webpageText and countwebpageText5 <= 15:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText5 += 1
                        
                        # =============================================================================
                        #                 #Add Payments Second window
                        # =============================================================================
                        driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[1]/button').click()
                        click_using_text(firstexposureText)
                        
                        countwebpageText6 = 0
                        while 'Cost Type' not in webpageText and countwebpageText6 <= 15:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText6 += 1
                            
                        if claimStatus == 'Closed':
                            paymentType = 'Final'
                        elif claimStatus == 'Open':
                            paymentType = 'Partial'
                        costType = row['Adjuster Cost Category']
                        paymentAmout = row['Grand Total']
                        
                        element_to_hover_over = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div/div/div[3]/div/div/div[1]/button')
                        actions = ActionChains(driver)
                        actions.move_to_element(element_to_hover_over)
                        actions.perform()
                        
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/div/label/div/div/div[1]/div[2]').click()
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/div/label/div/div/div[1]/div[2]/div/input').send_keys(str(costType),Keys.ENTER,Keys.TAB,
                        str(paymentAmout),Keys.ENTER,Keys.TAB,
                        str(paymentType),Keys.ENTER)
                        
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div/div[2]/div/div/div/div[4]/button[3]').click()
                        driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div/div[2]/div/div/div/div[4]/button[3]').click()
                        checkPointPayment = 0
                        countwebpageText6 = 0
                        while 'PAYMENT DETAILS' not in webpageText and countwebpageText6 <= 15:
                            time.sleep(1)
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText6 += 1
                        
                        time.sleep(3)
                        
                        if claimStatus == 'Closed':
                            Change_Claim_Status('Open','Closed')
                            
                            webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                            countwebpageText = 0
                            while 'Close Claim ' in webpageText and countwebpageText <= 60:
                                time.sleep(0.5)
                                webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                                countwebpageText += 1
                                
                            #close the claim    
                            driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').click()
                            
                            exposureClosCheckText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]').text
                            countexposureClosCheckText = 0
                            while 'CLOSED' not in exposureClosCheckText and countexposureClosCheckText <= 150:
                                time.sleep(0.5)
                                exposureClosCheckText = driver.find_element(by=By.XPATH,value='//*[@id="claim-page-wrapper"]/div/div[1]/div/div[1]').text
                                countexposureClosCheckText += 1    
                        
                        # =============================================================================
                        #  -------------------------> Closing the tasks <-----------------------------
                        # =============================================================================
                        counterCheckTaskButton = 0
                        while driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[1]/button[5]').get_attribute('title') != 'Tasks' and counterCheckTaskButton <= 20:
                            time.sleep(0.5)
                            counterCheckTaskButton += 1
                        
                        driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[1]/button[5]').click()
                        driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[1]/div/div[2]/div/button[2]').click()
                        # time.sleep(5)
                        if claimStatus == 'Closed':
                            # time.sleep(2)
                            checkReviewText = driver.find_element(by=By.XPATH,value='/html/body').text
                            checkReviewTextCounter = 0
                            while "Review Claim for Closure" not in checkReviewText and checkReviewTextCounter < 30:
                                time.sleep(1)
                                checkReviewText = driver.find_element(by=By.XPATH,value='/html/body').text
                                checkReviewTextCounter += 1
                                if checkReviewTextCounter == 29:
                                    print(checkReviewTextCounter, "Max wait exceeding while waiting for Review Claim Task")
                        taskListCountText = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[1]').text  
                        taskListCount = int(taskListCountText.replace('CURRENT(','').replace(')',''))
                        
                        t = 1
                        tCounter = 1
                        print('task numbers',taskListCount)
                        while t <= taskListCount:
                            print("Closing "+str(t)+" of "+str(taskListCount)+" tasks -->"+claimStatus)
                            if claimStatus == 'Closed':
                                # time.sleep(2)
                                checkReviewText = driver.find_element(by=By.XPATH,value='/html/body').text
                                checkReviewTextCounter = 0
                                while "Review Claim for Closure" not in checkReviewText and checkReviewTextCounter < 90:
                                    time.sleep(1)
                                    checkReviewText = driver.find_element(by=By.XPATH,value='/html/body').text
                                    checkReviewTextCounter += 1
                                    print(checkReviewTextCounter)
                                    if "Review Claim for Closure" in checkReviewText:
                                        print('YESS')
                                taskName = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/div').text
                                # print(claimStatus,'---->',taskName)
                                
                                if taskName == 'Reopen associated exposure(s) for further handling' or taskName == 'Review Claim for Closure':
                                    time.sleep(3)
                                    driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/div[2]/button').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/label/div').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/div[6]/button[2]').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/button').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[1]/div/div[2]/div/button[2]').click()
                                    print('closed the task','---->',taskName)              
                                    t -= 1
                                    taskListCount-=1
                            elif claimStatus == 'Open':
                                taskName = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/div').text
                                # print(taskName)
                                if taskName == 'Issue Payment' or taskName == 'Send Settlement letter':
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/button').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/label/div').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/div[6]/button[2]').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/button').click()
                                    driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[1]/div/div[2]/div/button[2]').click()
                        
                                    t -= 1
                                    taskListCount-=1
                            t += 1
                        
                        # print('done')
                        
                        
                        row['EOX Comments'] = 'Processed'
                        
                        # filename1 = str(filename).replace('.xlsx','')
                        # data.to_excel(LOCAL_OUTPUT+str(filename1)+'_completed.xlsx',index=False)
                        filename1_flag = 1
                        
                    elif claimStatus.upper() == 'CANCELLED':
                        row['EOX Comments'] = 'Claim State is CANCELLED'
                        
                    else:
                        row['EOX Comments'] = 'Unknown Error in Claim State'
                        
                    
                except Exception as e:
                    exception_flag = 1
                    try:
                        addComment = ''
                        
                        filename1 = str(filename).replace('.xlsx','')
                        data.to_excel(LOCAL_OUTPUT+str(filename1)+'_partially_completed.xlsx',index=False)
                        
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        # directory = LOCAL_ERROR+str(filenamedate)                     

                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        time.sleep(3)
                        if "Your password has expired. Please reset your password." in webpageText:
                            # send_text_email("Sagesure | RPA | Password Expired | "+str(usr))
                            log_error(fname, pos+1, exc_type, exc_obj, exc_tb, "Password expired!")
                            print("Password expired!")
                        
                        try:
                            if vendorCheckFlag == 0:
                                addComment = 'Check if Vendor is Present'
                        except:
                            pass
    
                        if 'We didn\'t find any results for the search' in webpageText:
                            log_error(fname, pos+1, exc_type, exc_obj, exc_tb, "We did not find any results for this Claim Number")
                            print("We didn\'t find any results for the search")
                            driver.save_screenshot(LOCAL_ERROR+fname+str(datetime.now().strftime('%m_%d_%Y %H:%M:%S'))+'\\'+str(pos)+'_'+str(clnumberfromexcel)+'_'+str(invoiceNumber)+'.png')
                            
                        else:
                            addCommentPageNotFound = ''
                            if 'Page not found' in webpageText:
                                addCommentPageNotFound = 'Page Not Found - '

                            if checkPointPayment == 1:
                                comments = addCommentPageNotFound + addComment + str((exc_tb.tb_lineno,exc_type,exc_obj))
                                log_error(fname, pos+1, exc_type, exc_obj, exc_tb, comments)
                                driver.save_screenshot(LOCAL_ERROR+fname+str(datetime.now().strftime('%m_%d_%Y %H:%M:%S'))+'\\'+str(pos)+'_'+str(clnumberfromexcel)+'_'+str(invoiceNumber)+'.png')
                                
                            elif checkPointPayment == 0:
                                comments = addCommentPageNotFound + addComment +'Completed the payment with an issue while closing tasks. Please check manually.'+ str((exc_tb.tb_lineno,exc_type,exc_obj))
                                log_error(fname, pos+1, exc_type, exc_obj, exc_tb, comments)
                                driver.save_screenshot(LOCAL_ERROR+fname+str(datetime.now().strftime('%m_%d_%Y %H:%M:%S'))+'\\'+str(pos)+'_'+str(clnumberfromexcel)+'_'+str(invoiceNumber)+'.png')
                                
                        # filename1 = str(filename).replace('.xlsx','')
                        # data.to_excel(LOCAL_OUTPUT+str(filename1)+'_completed.xlsx',index=False)
                        
                        try:
                            
                            if claimStatus == 'Closed':
                                driver.refresh()
                                claimSearchButton = '//*[@id="header"]/div/div[1]/div[1]/div'
                                wait_for_ele_presence(claimSearchButton,100)
                                Change_Claim_Status('Open','Closed')
                                yesCheckText = driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').text
                                countyesCheckText = 0
                                while yesCheckText != 'Yes' and countyesCheckText <= 20:
                                    yesCheckText = driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').text
                                    time.sleep(1)
                                    countyesCheckText += 1
                                driver.find_element(by=By.XPATH,value='/html/body/div[3]/div[2]/div[3]/button[2]').click()
                                
                                counterCheckTaskButton = 0
                                while driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[1]/button[5]').get_attribute('title') != 'Tasks' and counterCheckTaskButton <= 20:
                                    time.sleep(0.5)
                                    counterCheckTaskButton += 1
                                    
                                driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[1]/button[5]').click()
                                driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[1]/div/div[2]/div/button[2]').click()
                                # time.sleep(15)
                                if claimStatus == 'Closed':
                                    # time.sleep(2)
                                    checkReviewText = driver.find_element(by=By.XPATH,value='/html/body').text
                                    checkReviewTextCounter = 0
                                    while "Review Claim for Closure" not in checkReviewText and checkReviewTextCounter < 30:
                                        time.sleep(1)
                                        checkReviewText = driver.find_element(by=By.XPATH,value='/html/body').text
                                        checkReviewTextCounter += 1
                                        print(checkReviewTextCounter)
                                        if "Review Claim for Closure" in checkReviewText:
                                            print('YESS')
                                taskListCountText = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[1]').text  
                                taskListCount = int(taskListCountText.replace('CURRENT(','').replace(')',''))
                                # check_for_webpage_text('Review Claim for Closure',20)
                                t = 1
                                tCounter = 1
                                while t <= taskListCount:
                                    #print("Closing "+str(t)+" of "+str(taskListCount)+" tasks -->"+claimStatus)
                                    if claimStatus == 'Closed':
                                        
                                        taskName = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/div').text
                                        print(claimStatus,'---->',taskName)
                                        if taskName == 'Reopen associated exposure(s) for further handling' or taskName == 'Review Claim for Closure':
                                            # time.sleep(3)
                                            driver.find_element(by=By.XPATH,value='/html/body/div[2]/div/div[2]/div/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/div[2]/button').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/label/div').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/div[6]/button[2]').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/button').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[1]/div/div[2]/div/button[2]').click()
                                            print('closed the task','----->',taskName)                        
                                            t -= 1
                                            taskListCount-=1
                                    elif claimStatus == 'Open':
                                        taskName = driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/div').text
                                        # print(taskName)
                                        if taskName == 'Issue Payment' or taskName == 'Send Settlement letter':
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div[1]/div[2]/div['+str(t)+']/div/div[1]/button').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/label/div').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/div/div[6]/button[2]').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[2]/div/button').click()
                                            driver.find_element(by=By.XPATH,value='//*[@id="scaffold-wrapper"]/div/div[3]/div[5]/div/div[2]/div[1]/div/div[2]/div/button[2]').click()
                                
                                            t -= 1
                                            taskListCount-=1
                                    t += 1                   
                            claimStatus = '' 
                        except:
                            pass
                    except Exception as e:
                        try:
                            driver.quit()
                            pass
                        except:
                            pass
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        
                        # row['EOX Comments'] = "Closed unexpectedly. Please check manually for payment confirmation" + str((exc_tb.tb_lineno,exc_type,exc_obj))
                        log_error(fname, pos+1, exc_type, exc_obj, exc_tb, "Closed unexpectedly. Please check manually for payment confirmation")

                        driver = webdriver.Chrome(options=options)
                        driver.maximize_window()
                        driver.get(credfile.url)
                        driver.implicitly_wait(15)
                        
                        wait_for_ele_presence('//*[@id="app"]/div/form[2]/label/input')
                        driver.find_element(by=By.XPATH,value='//*[@id="app"]/div/form[2]/label/input').send_keys(usr,Keys.ENTER)
                        wait_for_ele_presence('//*[@id="app"]/div/form[2]/label[2]/input')
                        driver.find_element(by=By.XPATH,value='//*[@id="app"]/div/form[2]/label[2]/input').send_keys(pwd,Keys.ENTER)
                        time.sleep(3)
                        webpageText = driver.find_element(by=By.XPATH,value='/html/body').text
                        
                        if "Your password has expired. Please reset your password." in webpageText:
                            # send_text_email("Sagesure | RPA | Password Expired | "+str(usr))
                            log_error(fname, pos+1, exc_type, exc_obj, exc_tb, "Password expired!")
                            print("Password expired!")
                                
                        claimSearchButton = '//*[@id="header"]/div/div[1]/div[1]/div'
                        wait_for_ele_presence(claimSearchButton,100)
                        
                    
                    driver.refresh()
                    claimSearchButton = '//*[@id="header"]/div/div[1]/div[1]/div'
                    try:
                        wait_for_ele_presence(claimSearchButton,100)
                    except:
                        pass
                    vendorCheckFlag = ''
                    filename1 = str(filename).replace('.xlsx','')
                    data.to_excel(LOCAL_OUTPUT+str(filename1)+'_partially_completed.xlsx',index=False)
                        
                #print(pos+1)
                if pos % 10 == 0:
                    driver.refresh()
                    time.sleep(3)
                if pos % 50 == 0:
                    
                    driver.quit()
                    driver = webdriver.Chrome(options=options)
                    driver.maximize_window()
                    driver.get(credfile.url)
                    driver.implicitly_wait(15)
                    
                    wait_for_ele_presence('//*[@id="app"]/div/form[2]/label/input')
                    driver.find_element(by=By.XPATH,value='//*[@id="app"]/div/form[2]/label/input').send_keys(usr,Keys.ENTER)
                    wait_for_ele_presence('//*[@id="app"]/div/form[2]/label[2]/input')
                    driver.find_element(by=By.XPATH,value='//*[@id="app"]/div/form[2]/label[2]/input').send_keys(pwd,Keys.ENTER)
                    claimSearchButton = '//*[@id="header"]/div/div[1]/div[1]/div'
                    wait_for_ele_presence(claimSearchButton,100)

            if exception_flag == 1:
                pass
            else:
                filename1 = str(filename).replace('.xlsx','')
                data.to_excel(LOCAL_OUTPUT+str(filename1)+'_completed.xlsx',index=False)
            # driver.quit()
            # send_email_with_attachment(LOCAL_OUTPUT+str(filename1)+'_completed.xlsx')
        
        else:
            print(f"\nMissing headers: {missing_headers}")
            print(f"Processing stopped for file: {fName} due to column mismatch or missing headers")

            # Duration of File processing 
            endTime = time.time()
            print("Total Time taken for processing:",endTime-startTime)
            
            return fName

        # Duration of File processing
        endTime = time.time()
        print("Total Time taken for processing:",endTime-startTime)

        return filename1
    
    except:
        if filename1_flag:
            print("Error in processing file. Check file")

            # Duration of File processing
            endTime = time.time()
            print("Total Time taken for processing:",endTime-startTime)
            exc_type, exc_obj, exc_tb = sys.exc_info()
            log_error(fName, pos+1, exc_type, exc_obj, exc_tb, "Error while processing")
            return fName
        
        elif filename1_flag == 0:
            print("Error in processing. Check file and error_log")

            # Duration of File processing
            endTime = time.time()
            print("Total Time taken for processing:",endTime-startTime)
            exc_type, exc_obj, exc_tb = sys.exc_info()
            log_error(fName, pos+1, exc_type, exc_obj, exc_tb, "Chrome Crash")
            return fName

# call_process('test.xlsx',[('shriharim@eoxvantage.com','Gml9799/////////')])