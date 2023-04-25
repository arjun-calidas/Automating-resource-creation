from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import openpyxl
import pandas as pd
import os
import time
import re
from credentials import CREDENTIALS
from unidecode import unidecode

df = pd.read_excel('client.xlsx', sheet_name='Sheet1', header=0)
browser = webdriver.Firefox()
browser.maximize_window()
wait = WebDriverWait(browser, 10)

def get_to_login(browser, df):
    instance = df['instance'][0].lower()
    if instance == 'eu':
        browser.get('https://portal.myacolad.eu/account/login')
        cookie_button = wait.until(EC.element_to_be_clickable((By.ID, "consent")))
        cookie_button.click()
        email = CREDENTIALS['eu']['email']
        password = CREDENTIALS['eu']['password']
    elif instance in ['us', 'usa']:
        browser.get('https://portal.myacolad.com/account/login')
        cookie_button = wait.until(EC.element_to_be_clickable((By.ID, "consent")))
        cookie_button.click()
        email = CREDENTIALS['us']['email']
        password = CREDENTIALS['us']['password']

    elif instance in ['demo', 'DEMO', 'uat', 'UAT']:
        browser.get('https://portal-demo.myacolad.com/account/login')
        cookie_button = wait.until(EC.element_to_be_clickable((By.ID, "consent")))
        cookie_button.click()
        email = CREDENTIALS['demo']['email']
        password = CREDENTIALS['demo']['password']

    else:
        raise ValueError('Invalid instance specified - please check the excel sheet and add an instance')

    return email, password


def login_to_portal(browser, df):
    email, password = get_to_login(browser, df)
    email_input = browser.find_element("name",'Email')
    password_input = browser.find_element("name",'Password')
    email_input.send_keys(email)
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)
    return browser

login_to_portal(browser, df)

def navigate_to_clients_tab(browser):
    # Find the "Clients" tab element and click it
    wait = WebDriverWait(browser, 10)
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    clients_tab_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Clients")))
    clients_tab_link.click()

    # Wait for the "Clients" page to load
    wait.until(EC.title_contains("Clients"))

def click_add_client_button(browser):
    wait = WebDriverWait(browser, 10)
    add_client_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/admin/clients/client/0']")))
    add_client_button.click()

def add_client_name(browser, df):
    client_name = df['Client Name'][0]
    wait = WebDriverWait(browser, 10)
    client_name_input = wait.until(EC.element_to_be_clickable((By.ID, "ClientName")))
    client_name_input.clear()
    client_name_input.send_keys(client_name)


def enable_api_experience(browser):
    checkbox = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'IsPortalExperienceApiEnabled')))
    actions = ActionChains(browser)
    actions.move_to_element(checkbox).click().perform()
    # browser.execute_script("arguments[0].scrollIntoView();", checkbox)
    # checkbox.click()


def enter_tms_instance_id(browser, df):
    tms_id_input = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, 'PortalExperienceApiInstanceId')))
    actions = ActionChains(browser)
    actions.move_to_element(tms_id_input).click().send_keys(df['TMS instance id'].iloc[0]).perform()

def click_add_client_submit_button(browser):
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    add_client_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'submitForm')))
    add_client_button.click()


def search_client(browser, df):
    # Find the email field and paste the value from the DataFrame
    client_name = df['Client Name'].iloc[0]
    wait = WebDriverWait(browser, 10)
    client_name_input = wait.until(EC.element_to_be_clickable((By.ID, "table-search")))
    client_name_input.clear()
    client_name_input.send_keys(client_name)


def edit_client_settings(browser):
    settings_button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'far.fa-cog.fa-1-5x')))
    settings_button.click()

# modal_backdrop = (By.CLASS_NAME, 'modal-backdrop')


def click_add_source_language_button(browser):
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    add_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'button-add-source-language')))
    add_button.click()

def click_add_target_language_button(browser):
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    add_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'button-add-target-language')))
    add_button.click()

def click_add_service(browser):
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    add_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'button-add-service')))
    add_button.click()

def click_add_currency(browser):
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    add_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'button-add-currency')))
    add_button.click()

def click_add_file_category(browser):
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    add_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'button-add-file-category')))
    add_button.click()


def search_selection(browser, selection):
    wait = WebDriverWait(browser, 10)
    search_input = wait.until(EC.element_to_be_clickable((By.ID, "search-setting")))
    search_input.clear()
    search_input.send_keys(selection)

def search_new_api_service(browser):
    search_input = wait.until(EC.element_to_be_clickable((By.ID, "search-setting")))
    search_input.clear()
    search_input.send_keys('New API Service')
    search_input.send_keys(Keys.RETURN)
    time.sleep(1)

def click_checkboxes_for_services(browser):
    services = df['Services']
    if services.isnull().all():
        return
    services = services.dropna()
    for service in services:
        checkbox = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, f"//td[normalize-space()='{service}']/preceding-sibling::td/input[@type='checkbox']")))
        checkbox.click()

def click_selection_checkbox(browser):
    checkbox = browser.find_element('name', 'SettingId')
    if not checkbox.is_selected():
        checkbox.click()

def add_selection(browser):
    add_button = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'button-add-setting')))
    add_button.click()


def save_client_settings(browser):
    modal_backdrop = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'modal-backdrop')))
    button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "button-submit-settings")))
    button.click()
    

# Calling the functions to execute
#1. Create the client
navigate_to_clients_tab(browser)
click_add_client_button(browser)
add_client_name(browser, df)
enable_api_experience(browser)
enter_tms_instance_id(browser, df)
click_add_client_submit_button(browser)
# navigate_to_clients_tab(browser)

#2. Add the settings for the client
search_client(browser, df)
edit_client_settings(browser)

for selection in df['Source Languages'].dropna():
    click_add_source_language_button(browser)
    search_selection(browser, selection)
    click_selection_checkbox(browser)
    add_selection(browser)


for selection in df['Target Languages'].dropna():
    click_add_target_language_button(browser)
    search_selection(browser, selection)
    click_selection_checkbox(browser)
    add_selection(browser)

click_add_service(browser)
search_new_api_service(browser)
click_checkboxes_for_services(browser)
add_selection(browser)

for selection in df['Currency'].dropna():
    click_add_currency(browser)
    search_selection(browser, selection)
    click_selection_checkbox(browser)
    add_selection(browser)

if not df['File Category'].empty:
    for selection in df['File Category'].dropna():
        click_add_target_language_button(browser)
        search_selection(browser, selection)
        click_selection_checkbox(browser)
        add_selection(browser)
else:
    pass


save_client_settings(browser)








