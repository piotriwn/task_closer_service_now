import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By    
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import csv
import sys
import getpass
from datetime import datetime
import os

SN_URL = "service now URL"
FILTER_URL = r"service now filter URL"
PATH = r".\chromedriver.exe"

# log string
log = ""

# standard exception action
def standardException(ex, driver, msg = "Failed.\n", quit = False):
    global log
    line = msg
    line += str(ex) + "\n"
    log += line
    print(line)
    if (quit):
        driver.quit()
    return False

# print and append to log
def printAndLog(msg):
    global log
    log += msg
    print(msg)

# log in to Service Now
def logInFun(driver, login):
    driver.get(SN_URL)
    global log

    printAndLog("Logging in to Service Now... ")
    try:
        usernameBox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
        usernameBox.send_keys(login)
    except Exception as e:
        return standardException(e, driver, quit=True)

    for i in range(3): # 3 attempts to log in
        try:
            nextLoginPage = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "next")))
            nextLoginPage.click()
        except Exception as e:
            return standardException(e, driver, quit=True)

        try:
            externalLogin = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Use external login")))
            externalLogin.click()
            driver.implicitly_wait(1)

            submitButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login")))
            submitButton.click()
        except TimeoutException as ex:
            pass
        
        # search for company banner
        try:
            banner = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//a[@class="navbar-brand"]')))
            break
        except Exception as e:
            line = f"Failed to log-in in {i+1} attempt. Trying again...\n"
            log += line
            print(line)
        
        if (i==2):
            driver.quit()
            return False


    printAndLog("Success.\n")
    return True

# load CSV
def loadCSV():
    global log
    
    printAndLog("Loading CSV... ")
    tasks = []

    with open("tasks.csv", mode="r", encoding="utf-8-sig") as csv_File:
        try:
            csv_reader = csv.reader(csv_File, delimiter=";")
            for task in csv_reader:
                tasks.append(task)
        except Exception as e:
            line = "Failed.\n"
            line += str(e) + "\n"
            log += line
            print(line)
            return False   
        
    printAndLog("Success.\n")
    return tasks

# save logs to a file
def logToFile():
    filename =  datetime.now().strftime(r"%d_%m_%Y_%H_%M_%S") + "_" + getpass.getuser() + ".txt"
    pathname = os.path.abspath(os.path.dirname(__file__))    
    file_path = os.path.join(pathname , 'Logs' , filename)
    with open (file_path , "w") as f:
        f.write(log)

# go to filter page
def reachISApage(driver):
    global log
    driver.get(FILTER_URL)

    printAndLog("Attempting to reach filter page... ")
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[@id="task_breadcrumb"]//a//b[contains (text(), "{0}")]'.format("Short description contains ABC"))))
    except Exception as e:
        return standardException(e, driver,quit=True)

    printAndLog("Success.\n")
    return True

# search for task
def findTask(driver, task):
    global log
    
    printAndLog("Attempting to click on filter icon...")
    try:
        filterImg = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//a[@id="task_filter_toggle_image"]')))
        filterImg.click()
        printAndLog("Success.\n")
    except Exception as e:
        return standardException(e,driver)
    
    printAndLog("Attempting to gather elements inside filter... ")
    try:
        filterElements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="filterContainer"]//span')))
        printAndLog("Success.\n")
    except Exception as e:
        return standardException(e,driver)
    
    printAndLog("Attempting to add keyword element if not present... ")
    filterElementsText = []
    for item in filterElements:
        try:
            filterElementsText.append(item.text)
        except:
            pass
    if (filterElementsText == []):
        printAndLog("Error parsing filter elements.\n")
        # driver.quit()
        return False

    if ("Keywords" not in filterElementsText):
        printAndLog("Keywords not present. Attempting to add keyword condition... ")
        try:
            time.sleep(3)
            addNew = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Add a new AND filter condition"]')))
            addNew.click()
            time.sleep(3)
            addNewSelection = driver.find_elements_by_xpath('//select[@aria-label="Choose Field"]')[-1]
            addNewSelection = Select(addNewSelection)
            addNewSelection.select_by_visible_text("Keywords")
            time.sleep(3)
            printAndLog("Success.\n")
        except Exception as e:
            return standardException(e,driver)
    
    printAndLog("Attempting to search for the task... ")
    try:
        keywordWE = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//tr[@class="filter_row_condition"]//input[@aria-label="Input value"]')))
        keywordWE.clear()
        keywordWE.send_keys(task)
        runButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//button[@id="test_filter_action_toolbar_run"]')))
        runButton.click()
        printAndLog("Success.\n")
    except Exception as e:
        return standardException(e,driver)

    return True
   
# check if closed
def checkClosed(driver, task):
    global log

    printAndLog("Attempting to check if the task has already been closed...")
    try:
        state = driver.find_element_by_xpath('//*[@id="sc_task.state"]//option[@selected="SELECTED"]')
        printAndLog("Status has been retrieved.\n")
    except Exception as e:
        return standardException(e,driver)

    if (state.text == "Closed Complete"):
        printAndLog(f"-----> Ticket {task} has already closed.\n")
        return "Closed"
    else:
        printAndLog(f"Ticket {task} has not been closed.\n")
        return "Not closed"

# open the task
def openTask(driver):
    global log
    printAndLog("Attempting to open the task... ")
    try:
        taskLink = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//a[@class="linked formlink"]')))
        taskLink.click()
        printAndLog("Success.\n")
        return True
    except Exception as e:
        return standardException(e,driver)

# close task
def closeTask(driver, username):
    global log

    saveButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sysverb_update_and_stay"]')))

    printAndLog("Attempting to assign the ticket... ")
    try:
        assignedTo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sys_display.sc_task.assigned_to"]')))
        assignedTo.clear()
        assignedTo.send_keys(username)
        time.sleep(2)
        assignedTo.send_keys(Keys.TAB)
        time.sleep(2)
        saveButton.click()
        time.sleep(2)
        assignedTo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sys_display.sc_task.assigned_to"]')))
        if (assignedTo.get_attribute("value").lower() != username.lower() ):
            raise Exception("Ticket has not been assigned successfully to a proper person.")
        else:
            printAndLog("Success.\n")
    except Exception as e:
        return standardException(e,driver)

    printAndLog("Attempting to paste a worknote... ")
    try:
        worknoteElement = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "activity-stream-work_notes-textarea")))
        worknoteElement.send_keys("As requested.")
        time.sleep(3)
        saveButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sysverb_update_and_stay"]')))
        saveButton.click()
        time.sleep(5)
        streamElements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//li[@class="h-card h-card_md h-card_comments"]//span[@class="sn-widget-textblock-body sn-widget-textblock-body_formatted"]')))
        worknoteText = []
        for item in streamElements:
            try:
                worknoteText.append(item.text)
            except:
                pass
        if ("As requested." not in worknoteText):
            raise Exception("Worknote has not been added successfully.")
        else:
            printAndLog("Success.\n")
    except Exception as e:
        return standardException(e,driver)
    
    printAndLog("Attempting to change the status of the ticket... ")
    try:
        state = Select(driver.find_element_by_xpath('//*[@id="sc_task.state"]'))
        state.select_by_value("3")
        time.sleep(2)
        saveButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sysverb_update_and_stay"]')))
        saveButton.click()
        state = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//select[@id="sc_task.state"]//option[@selected="SELECTED"]')))
        if (state.text != "Closed Complete"):
            raise Exception("State has not been changed successfully.")
        else:
            printAndLog("Success.\n")
    except Exception as e:
        return standardException(e,driver)
    
    return True

# main function
def main():
    global log
    driver = webdriver.Chrome(PATH)
    login = str(sys.argv[1])
    usern = str(sys.argv[2])

    logInResult = logInFun(driver, login)
    tasks = loadCSV()
    if (not logInResult  or not tasks):
        logToFile()
        sys.exit()

    for task in tasks:
        printAndLog(f"\n------------------------\nParsing task {task}.\n")

        reachedISA = reachISApage(driver)
        if (not reachedISA):
            logToFile()
            sys.exit()

        taskFound = findTask(driver, task)
        if (not taskFound):
            printAndLog(f"Error parsing task {task}.\n")
            continue

        openedTask = openTask(driver)
        if (not openedTask):
            printAndLog(f"Error opening {task}.\n")
            continue

        checkIfClosed = checkClosed(driver, task)
        if (checkIfClosed == "Not closed"):
            pass
        else:
            continue

        closedTask = closeTask(driver, usern)
        if (not closedTask):
            printAndLog(f"Error closing {task}.\n")
            continue
            
        printAndLog(f"-----> {task} has been closed SUCCESSFULLY.\n")
    
    logToFile()

if __name__=="__main__":
    main()

