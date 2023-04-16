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
    """

    # Open url in incognito | reference: https://www.tutorialspoint.com/how-to-open-browser-window-in-incognito-private-mode-using-python-selenium-webdriver 
    c = webdriver.ChromeOptions()
    c.add_argument("--incognito")
    driver = webdriver.Chrome('chromedriver', options=c)
    driver.implicitly_wait(0.5)
    driver.get(url)

    # Log in (fill in username, click button; repeat for password)
    driver.find_element(By.ID, "i0116").send_keys(username) # references: https://www.thepythoncode.com/article/automate-login-to-websites-using-selenium-in-python, https://pythonbasics.org/selenium-keyboard/  
    time.sleep(1)
    driver.find_element(By.ID, "i0116").send_keys(Keys.ENTER)  # reference: https://pythonbasics.org/selenium-keyboard/ 
    time.sleep(1)
    driver.find_element(By.ID, "i0118").send_keys(password)
    time.sleep(1)
    driver.find_element(By.ID, "i0118").send_keys(Keys.ENTER)
    time.sleep(1)

    # Wait for login process to complete
    WebDriverWait(driver=driver, timeout=10).until(
        lambda x: x.execute_script("return document.readyState === 'complete'"))
    
    try:
        driver.find_element(By.XPATH, "//input[@id='idSIButton9' and @value='Yes']").click()
    except:
        print("[!] Login failed. Incorrect username or password used.")
        return ValueError
    
    # Wait for login process to complete
    WebDriverWait(driver=driver, timeout=100).until(
        lambda x: x.execute_script("return document.readyState === 'complete'"))
    
    time.sleep(50)

    # try:
    #     driver.find_element(By.XPATH, "//button[@value='Yes']").click()
    # except:
    #     print("[!] Login failed. Incorrect username or password used.")
    #     return ValueError

    # # Wait for login process to complete
    # WebDriverWait(driver=driver, timeout=10).until(
    #     lambda x: x.execute_script("return document.readyState === 'complete'"))
    # sucess_message = "Don't show this again"

    # # Retrieve any errors found.
    # sucesses = driver.find_elements(By.CLASS_NAME, "col-md-24 form-group checkbox")

    # # When errors are found, the login will fail
    # if any(sucess_message in e.text for e in sucesses):
    #     print("[+] Login successful")
    # else:
    #     print("[!] Login failed")

    driver.implicitly_wait(50)
    # driver.close()

def check_login(driver):
    # Wait for login process to complete
    WebDriverWait(driver=driver, timeout=10).until(
        lambda x: x.execute_script("return document.readyState === 'complete'"))
    error_message = "Incorrect username or password."

    # Retrieve any errors found.
    errors = driver.find_elements(By.ID, "passwordError")

    # When errors are found, the login will fail
    if any(error_message in e.text for e in errors):
        print("[!] Login failed")
    else:
        print("[+] Login successful")

    
    # Wait for login process to complete
    WebDriverWait(driver=driver, timeout=100).until(
        lambda x: x.execute_script("return document.readyState === 'complete'"))
    error_message = "Forgot my password"

    # Retrieve any errors found.
    errors = driver.find_elements(By.ID, "idA_PWD_ForgotPassword")
    print(errors)

    # When errors are found, the login will fail
    if any(error_message in e.text for e in errors):
        print("[!] Login failed")
    else:
        print("[+] Login successful")

    driver.implicitly_wait(50)
    # driver.close()




username = input("Enter your Babson email -> ")
password = input("Enter your password -> ")
website_login(
    "https://secure.signin.spglobal.com/sso/saml2/0oa1mqx8p77XSX10T1d8/app/spglobaliam_sp_1/exk1mregn1oWwP2NB1d8/sso/saml?RelayState=https://www.capitaliq.com/CIQDotNet/saml-sso.aspx", username, password)
