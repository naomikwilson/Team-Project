from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

def website_login(url, username, password):
    """
    Logs user into website (url) using username and password

    Sources for code:
    - https://www.thepythoncode.com/article/automate-login-to-websites-using-selenium-in-python 
    - https://www.tutorialspoint.com/how-to-open-browser-window-in-incognito-private-mode-using-python-selenium-webdriver
    - https://www.geeksforgeeks.org/find_element_by_id-driver-method-selenium-python/ 
    - https://pythonbasics.org/selenium-keyboard/ 
    - https://selenium-python.readthedocs.io/locating-elements.html 
    - ChatGPT aided in the debugging process
    """

    # Open url in incognito
    c = webdriver.ChromeOptions()
    c.add_argument("--incognito")
    driver = webdriver.Chrome('chromedriver', options=c)
    driver.implicitly_wait(0.5)
    driver.get(url)

    # Log in (fill in username, click button; repeat for password)
    driver.find_element(By.ID, "i0116").send_keys(username)
    time.sleep(1)
    driver.find_element(By.ID, "i0116").send_keys(Keys.ENTER)
    time.sleep(1)
    driver.find_element(By.ID, "i0118").send_keys(password)
    time.sleep(1)
    driver.find_element(By.ID, "i0118").send_keys(Keys.ENTER)
    time.sleep(1)
    login_wait(driver)

    # Wait for login process to complete
    # WebDriverWait(driver=driver, timeout=10).until(
    #     lambda x: x.execute_script("return document.readyState === 'complete'"))
    # # time.sleep(1)
    
    try:
        driver.find_element(By.XPATH, "//input[@id='idSIButton9' and @value='Yes']").click()
    except:
        print("[!] Login failed. Incorrect username or password used.")
        return ValueError
    time.sleep(1)

    login_wait(driver)

    # WebDriverWait(driver=driver, timeout=30).until(
    #     lambda x: x.execute_script("return document.readyState === 'complete'"))
    time.sleep(10)
    
    # driver.implicitly_wait(500)
    driver.close()

def login_wait(driver):
    WebDriverWait(driver=driver, timeout=30).until(
        lambda x: x.execute_script("return document.readyState === 'complete'"))
    
def capital_IQ_excel():
    driver = webdriver.Chrome('chromedriver')
    driver.get("https://secure.signin.spglobal.com/sso/saml2/0oa1mqx8p77XSX10T1d8/app/spglobaliam_sp_1/exk1mregn1oWwP2NB1d8/sso/saml?RelayState=https://www.capitaliq.com/CIQDotNet/saml-sso.aspx")
    
    # login_wait(driver)
    # time.sleep(15)
    # driver.find_element(By.ID, "SearchTopBar").send_keys(company)
    # driver.find_element(By.ID, "SearchTopBar").send_keys(Keys.ENTER)

    driver.find_element(By.CLASS_NAME, "cSearchBoxDisabled").click()
    driver.find_element(By.CLASS_NAME, "cSearchBox").send_keys(company)
    driver.find_element(By.CLASS_NAME, "cSearchBox").send_keys(Keys.ENTER)

    login_wait(driver)
    time.sleep(15)
    driver.close()

def login_wait(driver):
    WebDriverWait(driver=driver, timeout=30).until(
        lambda x: x.execute_script("return document.readyState === 'complete'"))

username = input("Enter your Babson email -> ")
password = input("Enter your password -> ")
company = input("Enter a company -> ")
website_login(
    "https://secure.signin.spglobal.com/sso/saml2/0oa1mqx8p77XSX10T1d8/app/spglobaliam_sp_1/exk1mregn1oWwP2NB1d8/sso/saml?RelayState=https://www.capitaliq.com/CIQDotNet/saml-sso.aspx", username, password)
