from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import shutil
import re

def create_driver():
    """
    Creates driver that opens files in incognito.

    Source for code:
    - https://www.tutorialspoint.com/how-to-open-browser-window-in-incognito-private-mode-using-python-selenium-webdriver 
    """
    c = webdriver.ChromeOptions()
    c.add_argument("--incognito")
    driver = webdriver.Chrome('chromedriver', options=c)
    return driver

def login(driver, url, username, password):
    """
    Logs user into website (url) using username and password.

    Sources for code:
    - https://www.thepythoncode.com/article/automate-login-to-websites-using-selenium-in-python 
    - https://www.geeksforgeeks.org/find_element_by_id-driver-method-selenium-python/ 
    - https://pythonbasics.org/selenium-keyboard/ 
    - ChatGPT aided in the debugging process
    """

    driver.implicitly_wait(0.5)
    driver.get(url)

    # Log in (fill in username, click button; repeat for password)
    driver.find_element(By.ID, "i0116").send_keys(username)
    driver.find_element(By.ID, "i0116").send_keys(Keys.ENTER)
    time.sleep(1)
    driver.find_element(By.ID, "i0118").send_keys(password)
    driver.find_element(By.ID, "i0118").send_keys(Keys.ENTER)
    driver_wait(driver, 1)
    # driver.implicitly_wait(5)
    
    # Tries to find and click on "yes" button on "stay signed in?" page
    # If this is not possible, it means an error has occured (username or password is incorrect)
    try:
        driver.find_element(By.XPATH, "//input[@id='idSIButton9' and @value='Yes']").click()
    except:
        print("[!] Login failed. Incorrect username or password used.")
        return ValueError
    driver_wait(driver, 2)
    # driver.implicitly_wait(10)
    # time.sleep(10)

def capital_IQ(driver, company):
    """
    Downloads Comparative Analysis file for the company.

    Returns modified company name (to be used for move_file() function).
    
    Sources for code (in addition to sources for login() function):
    - https://selenium-python.readthedocs.io/locating-elements.html 
    - https://stackoverflow.com/questions/14596884/remove-text-between-and 
    """
    # Puts user-inputted company name into search box and clicks on first result on page
    driver.find_element(By.ID, "onetrust-accept-btn-handler").click()
    driver.find_element(By.CLASS_NAME, "cSearchBoxDisabled").click()
    driver.find_element(By.CLASS_NAME, "cSearchBox").send_keys(company)
    driver.find_element(By.CLASS_NAME, "cSearchBox").send_keys(Keys.ENTER)

    driver.find_element(By.XPATH, "//tr[@id='SR0']/td[@class='NameCell']/div/span/a").click()
    driver.find_element(By.ID, "ll_7_26_2305").click()

    # obtain the modified company name on page (to be used for locating and moving file later)
    company_name = driver.find_element(By.XPATH, "//span[@id='_CPH__companyPageTitle_PageHeaderLabel']/span[1]").text
    company_name = company_name.replace(",", "")
    company_name = company_name.replace(".", "")
    company_name = re.sub("[\(\[].*?[\)\]]", "", company_name)

    # download Excel file
    driver.find_element(By.XPATH, "//img[@title='Download Comp Set to Excel']").click()

    driver_wait(driver, 1)
    time.sleep(15)
    driver.close()

    print(f"{company_name}")
    return company_name

def move_file(user, company_name):
    """
    Moves Excel file from Downloads folder to excel_files folder in this repository.

    Sources for code:
    - https://www.learndatasci.com/solutions/python-move-file/
    """
    old_path = f"C:/Users/{user}/Downloads/Company Comparable Analysis {company_name}.xls"
    new_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    shutil.move(old_path, new_path)

def driver_wait(driver, action):
    """
    Driver waits util action can be done.

    actions:
    - 1: wait for page to load or download to complete
    - 2: wait for Capital IQ page to load

    Source for code: 
    - https://www.thepythoncode.com/article/automate-login-to-websites-using-selenium-in-python  
    """
    if action == 1:
        WebDriverWait(driver=driver, timeout=15).until(
            lambda x: x.execute_script("return document.readyState === 'complete'"))
    else:
        WebDriverWait(driver=driver, timeout=30).until(
            lambda x: x.find_element(By.ID, "onetrust-accept-btn-handler"))
    
def main():
    username = input("Enter your Babson email -> ")
    password = input("Enter your password -> ")
    company = input("Enter a company -> ")
    user = input("Enter your Windows user name -> ")
    driver = create_driver()
    login(driver, "https://secure.signin.spglobal.com/sso/saml2/0oa1mqx8p77XSX10T1d8/app/spglobaliam_sp_1/exk1mregn1oWwP2NB1d8/sso/saml?RelayState=https://www.capitaliq.com/CIQDotNet/saml-sso.aspx", username, password)
    company_name = capital_IQ(driver, company)
    move_file(user, company_name)

if __name__ == '__main__':
    main()