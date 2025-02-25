import pickle
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

OUTLOOK_URL = "https://outlook.office365.com/calendar/view/workweek"
EMAIL = "username@uwaterloo.ca"
PASSWORD = "password"
AUTH_TIME_LIMIT = 300 # in seconds

def login_store_cookies():
    #initialize webdriver and open URL
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(OUTLOOK_URL)

    #wait for login page and enter email
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "loginfmt"))
    )
    email_field = driver.find_element(By.NAME, "loginfmt")
    email_field.send_keys(EMAIL)
    email_field.send_keys(Keys.RETURN)

    #wait for redirected UW page and enter password
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "passwordInput"))
    )
    password_field = driver.find_element(By.ID, "passwordInput")
    password_field.send_keys(PASSWORD)
    password_field.send_keys(Keys.RETURN)

    #wait for DUO authentication page
    try:
        #wait to click trust device
        WebDriverWait(driver, AUTH_TIME_LIMIT).until(
            EC.presence_of_element_located((By.ID, "trust-browser-button"))
        )
        trust_btn = driver.find_element(By.ID, "trust-browser-button")
        trust_btn.click()

        #wait to click Sign-In
        WebDriverWait(driver, AUTH_TIME_LIMIT).until(
            EC.presence_of_element_located((By.ID, "idSIButton9"))
        )
        cnt_btn = driver.find_element(By.ID, "idSIButton9")
        cnt_btn.click()
    #>1 minute has passed, inform of missing DUO authentication
    except Exception as e:
        print(f"Error during Duo authentication: {e}")
        driver.quit()
        return

    #wait for successful login to Outlook
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Mail']"))
    )

    #save cookies to file
    cookies = driver.get_cookies()
    with open("cookies.pkl", "wb") as file:
        pickle.dump(cookies, file)

    time.sleep(30)
    driver.quit()

login_store_cookies()
