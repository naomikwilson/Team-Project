from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By


def website_login(url, username, password):

    driver = webdriver.Chrome('chromedriver')

    driver.get(url)

    # find username field and send username to the input field
    driver.find_element("id", "username").send_keys(username)

    # find password field and insert password
    driver.find_element("id" "password").send_keys(password)

    # find login button
    driver.find_element("id", "myLoginButton").click()

    # Wait for login process to complete.
    WebDriverWait(driver=driver, timeout=10).until(
        lambda x: x.execute_script("return document.readyState === 'complete'"))
    # Verify that the login was successful.
    error_message = "Incorrect username or password."
    # Retrieve any errors found.
    errors = driver.find_elements(By.CLASS_NAME, "flash-error")

    # When errors are found, the login will fail.
    if any(error_message in e.text for e in errors):
        print("[!] Login failed")
    else:
        print("[+] Login successful")

    driver.implicitly_wait(10)
    driver.close()


username = input("Enter username -> ")
password = input("Enter your password -> ")
website_login(
    "https://www.capitaliq.com/CIQDotNet/my/dashboard.aspx", username, password)
