import google.generativeai as genai 
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui 
import os
import dotenv
import time
import pyperclip

dotenv.load_dotenv()

genai.configure(api_key=os.environ["apikey"])
model = genai.GenerativeModel("gemini-1.5-flash-002")


while True:
    assignment_work = input("Describe and tell about your assignment : ")
    response = model.generate_content(assignment_work).text
    pyperclip.copy(response)

    #login into the account
    driver = webdriver.Chrome()
    driver.get("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=4765445b-32c6-49b0-83e6-1d93765276ca&redirect_uri=https%3A%2F%2Fwww.office.com%2Flandingv2&response_type=code%20id_token&scope=openid%20profile%20https%3A%2F%2Fwww.office.com%2Fv2%2FOfficeHome.All&response_mode=form_post&nonce=638697079660232223.MjAzNmJhNWEtNTQ3Zi00MjEyLTk3MGQtNTk1NDQ3NDI4ZjYxMGEyOWE5Y2MtMTFjMi00M2U4LWI3YzItYWM5MGY0MjM0MjM5&ui_locales=en-US&mkt=en-US&client-request-id=12cca6bb-ba59-4a64-ad64-f8433086721a&state=6tfmWUMRyrWx5vC6FxcpOHqS75ooLwbWoGb_hvXIIVwEdiwTDNd8sACc-hJn8o-u0wLc9eVqkDgZmCci4Vo6DSW0W3Pfzbj2U8sgZM5oTy_YS0jp2Rz-71umA8TPS1DEnMdEuFYIoezWN-9i2K78xsPZh9LxG6VWyfTYIO0kCwCD69IlEblaEM2N3EnBk-5M8z4jqEr6ahCcBaZbHu6KGZHPKLtMnZP6ZTqZ6A2K6Jy5qzFED0h4MQWE42Mw2ets8WigO0-BenfGT6XklIvRObczwEEfRSh8sqf6hwOEkV0&x-client-SKU=ID_NET8_0&x-client-ver=7.5.1.0&sso_reload=true")
    time.sleep(3)

    emailinput = driver.find_element(By.XPATH, './/input[@class="form-control ltr_override input ext-input text-box ext-text-box"]').click()
    pyautogui.typewrite(os.environ["email"])
    pyautogui.press('enter')
    time.sleep(3)

    passwordinput = driver.find_element(By.NAME, "passwd")
    pyautogui.typewrite(os.environ['password'])
    pyautogui.press('enter')
    time.sleep(4)

    #ask for stay to login 
    pyautogui.press('enter')
    time.sleep(5)

    #open word sheet and write on it 
    new_word_doc = driver.find_element(By.CLASS_NAME, "CreateTilesListControl-module__create-tiles-list__item-refresh__KDk4e")
    new_word_doc.click()
    time.sleep(20)

    #finally writing assignment
    pyautogui.hotkey("ctrl","v")

    user_exit = input("Do you want to exit the program y/n : ")

    if user_exit == "y":
        driver.quit()
        break