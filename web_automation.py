from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import shutil
import glob
import os
import os.path


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

    # Log in (fill in username, click button; repeat for password).
    driver.find_element(By.ID, "i0116").send_keys(username)
    driver.find_element(By.ID, "i0116").send_keys(Keys.ENTER)
    time.sleep(1)
    driver.find_element(By.ID, "i0118").send_keys(password)
    driver.find_element(By.ID, "i0118").send_keys(Keys.ENTER)
    driver_wait(driver, 1)

    # Try to find and click on "yes" button on "stay signed in?" page.
    # If this is not possible, it means an error has occured (username or password is incorrect).
    try:
        driver.find_element(
            By.XPATH, "//input[@id='idSIButton9' and @value='Yes']").click()
    except:
        raise ValueError(
            "[!] Login failed. Incorrect username or password used.")
    driver_wait(driver, 2)


def capital_IQ(driver, companies):
    """
    Downloads Comparative Analysis file for the company.

    Source for code (in addition to sources for login() function):
    - https://selenium-python.readthedocs.io/locating-elements.html 
    """
    # Click "Accept Cookies".
    driver.find_element(By.ID, "onetrust-accept-btn-handler").click()

    # Put user-inputted company names into search box and click on first result on page.
    companies = companies.replace(" ", "")
    companies_list = companies.split(",")

    for names in companies_list:
        driver.find_element(By.CLASS_NAME, "cSearchBoxDisabled").click()
        driver.find_element(By.CLASS_NAME, "cSearchBox").send_keys(names)
        driver.find_element(By.CLASS_NAME, "cSearchBox").send_keys(Keys.ENTER)

        driver.find_element(
            By.XPATH, "//tr[@id='SR0']/td[@class='NameCell']/div/span/a").click()
        driver.find_element(By.ID, "ll_7_26_2305").click()

        # Download Excel file.
        driver.find_element(
            By.XPATH, "//img[@title='Download Comp Set to Excel']").click()
        time.sleep(2)

    driver.close()

    # Company_list will be used in move_file() function to get number of files that need to be moved.
    return companies_list


def move_file(username, companies_list):
    """
    Moves Excel files from Downloads folder to excel_files folder in this repository.

    Sources for code:
    - https://www.learndatasci.com/solutions/python-move-file/
    - https://datatofish.com/latest-file-python/ 
    - ChatGPT aided in the debugging process
    """

    user, __ = username.split("@")
    number_of_files = len(companies_list)

    folder_path = f"C:/Users/{user}/Downloads/"
    file_type = "/*xls"
    files = glob.glob(folder_path + file_type)

    # Sort files based on creation time, select number of company files downloaded.
    recent_files = sorted(files, key=os.path.getctime,
                          reverse=True)[:number_of_files]

    # Move files to excel_files folder in this repository.
    excel_folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    for file in recent_files:
        shutil.move(file, excel_folder_path)

    print("Download complete. File is now in the excel_files folder.")


def driver_wait(driver, action):
    """
    Driver waits util action can be done.

    Actions:
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
    companies = input(
        "Enter companies of interest; separate each by a comma -> ")
    driver = create_driver()
    login(driver, "https://secure.signin.spglobal.com/sso/saml2/0oa1mqx8p77XSX10T1d8/app/spglobaliam_sp_1/exk1mregn1oWwP2NB1d8/sso/saml?RelayState=https://www.capitaliq.com/CIQDotNet/saml-sso.aspx", username, password)
    try:
        companies_list = capital_IQ(driver, companies)
        move_file(username, companies_list)
    except:
        raise ValueError(
            "[!] Incorrect company name or format; if entering multiple companies, please include commas between company names.")


if __name__ == '__main__':
    main()
