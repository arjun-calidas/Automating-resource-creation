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

df = pd.read_excel('users.xlsx', sheet_name='Sheet1', header=0)
browser = webdriver.Firefox()
# browser = webdriver.Chrome()
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


def navigate_to_users_tab(browser):
    # Find the "Users" tab element and click it
    wait = WebDriverWait(browser, 10)
    users_tab_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Users")))
    users_tab_link.click()

    # Wait for the "Users" page to load
    wait.until(EC.title_contains("Users"))

navigate_to_users_tab(browser)

def click_add_user_button(browser):
    wait = WebDriverWait(browser, 10)
    add_user_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-success")))
    add_user_button.click()

    # Wait for the "Add User" page to load
    wait.until(EC.title_contains("Add User"))



def add_user_details(browser, excel_file_name, row_index):


    row = df.loc[row_index]

    # Extract data from row and remove any special characters for data copied from the ppt
    email = unidecode(row['Email']).strip()
    first_name = unidecode(row['First Name']).strip()
    last_name = unidecode(row['Last Name']).strip()
    password = row['Password']

    # Fill in user details
    email_input = wait.until(EC.presence_of_element_located((By.ID, "Email")))
    email_input.clear()
    email_input.send_keys(email)

    first_name_input = wait.until(EC.presence_of_element_located((By.ID, "FirstName")))
    first_name_input.clear()
    first_name_input.send_keys(first_name)

    last_name_input = wait.until(EC.presence_of_element_located((By.ID, "LastName")))
    last_name_input.clear()
    last_name_input.send_keys(last_name)

    password_input = wait.until(EC.presence_of_element_located((By.ID, "Password")))
    password_input.clear()
    password_input.send_keys(password)

    confirm_password_input = wait.until(EC.presence_of_element_located((By.ID, "ConfirmPassword")))
    confirm_password_input.clear()
    confirm_password_input.send_keys(password)



def select_timezone(browser, timezone):
    # Define the wait variable
    wait = WebDriverWait(browser, 10)
    # Select the desired timezone option
    timezone_dropdown = wait.until(EC.visibility_of_element_located((By.ID, "TimezoneId")))
    timezone_dropdown.click()
    brussels_option = wait.until(EC.visibility_of_element_located((By.XPATH, "//option[text()='(UTC+01:00) Brussels, Copenhagen, Madrid, Paris']")))
    brussels_option.click()



def select_language(browser):
    # Define the wait variable
    wait = WebDriverWait(browser, 10)
    # Select the desired language option
    language_dropdown = wait.until(EC.visibility_of_element_located((By.ID, "Language")))
    language_dropdown.click()
    # english_us_option = wait.until(EC.visibility_of_element_located((By.ID, "english-us-option")))
    english_us_option = wait.until(EC.visibility_of_element_located((By.XPATH, "//option[text()='English (en-US)']")))
    english_us_option.click()



def select_role(browser, role):
    wait = WebDriverWait(browser, 10)
    # instance = df['instance'][0].lower()
    if role.lower() == 'client':
        role_radio_button = wait.until(EC.element_to_be_clickable((By.ID, "Client")))
    elif role.lower() == 'supplier':
        role_radio_button = wait.until(EC.element_to_be_clickable((By.ID, "Supplier")))
    else:
        raise ValueError("Invalid role specified - please check the excel sheet and add a role")
    
    role_radio_button.click()




def select_client(browser, excel_file_name, row_index):
    client_name = df.loc[row_index, 'Client Name']

    # Find the client dropdown element and click it
    wait = WebDriverWait(browser, 10)
    client_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ClientId")))
    client_dropdown.click()

    # Find the client option element with the specified text and click it
    client_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//option[text()='{client_name}']")))
    client_option.click()



def select_supplier(browser):
    # Wait for the select element to be visible
    select_element = WebDriverWait(browser, 10).until(
        EC.visibility_of_element_located((By.ID, "SupplierId"))
    )

    # Click on the select element to open the dropdown
    select_element.click()

    # Wait for the Acolad option to be visible and then click on it
    acolad_option = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, "//option[text()='Acolad']")))
    acolad_option.click()




def additional_roles(browser, df, row_index):
    parent_div = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "role-client-options")))
    div_ul = parent_div.find_element(By.CLASS_NAME, "select2-selection__rendered")


    # Get the additional roles for the specified user
    add_roles = df.iloc[row_index]['additional roles']

    if pd.isna(add_roles):
        return  # Return if there are no additional roles for this user

    # Iterate over the roles to add and select each one in the dropdown
    for role in re.split(',|\n', add_roles):
        # Select the role option by its text value
        browser.execute_script("arguments[0].scrollIntoView();", div_ul)
        div_ul.click()
        search_box = parent_div.find_element(By.CLASS_NAME, 'select2-search__field')

        # Type the role name into the search box and press Enter to select it
        search_box.send_keys(role.strip())
        search_box.send_keys(Keys.RETURN)
        
        # time.sleep(1)  # Wait for dropdown to load


def additional_roles_edit_user(browser, df, row_index):
    div_span = browser.find_element(By.XPATH, "//span[contains(@class, 'select2-selection--multiple')]")

    # Get the additional roles for the specified user
    add_roles = df.iloc[row_index]['additional roles']

    if pd.isna(add_roles):
        return  # Return if there are no additional roles for this user

    # Iterate over the roles to add and select each one in the dropdown
    for role in re.split(',|\n', add_roles):
        # Select the role option by its text value
        
        browser.execute_script("arguments[0].scrollIntoView();", div_span)
        div_span.click()
        search_box = browser.find_element(By.XPATH, "//input[contains(@class, 'select2-search__field')]")
        

        # Type the role name into the search box and press Enter to select it
        search_box.send_keys(role.strip())
        search_box.send_keys(Keys.RETURN)


def click_activated_radio_button(browser):
    wait = WebDriverWait(browser, 10)
    activated_radio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='SendActivationEmail'][@value='false']")))
    activated_radio_button.click()


def click_create_user_button(browser):
    wait = WebDriverWait(browser, 10)
    create_user_button = wait.until(EC.element_to_be_clickable((By.ID, "button-create-user")))
    create_user_button.click()

def click_save_user_button(browser):
    wait = WebDriverWait(browser, 10)
    create_user_button = wait.until(EC.element_to_be_clickable((By.ID, "button-save-user")))
    create_user_button.click()

def add_timestamp(df):
    now = datetime.datetime.now()
    df['timestamp'] = now
    return df


def search_user(browser, df, row_index):
    # Find the email field and paste the value from the DataFrame
    email = df.loc[row_index, 'Email']
    wait = WebDriverWait(browser, 10)
    email_input = wait.until(EC.element_to_be_clickable((By.ID, "table-search")))
    email_input.clear()
    email_input.send_keys(email)

def edit_user(browser):
    wait = WebDriverWait(browser, 10)
    edit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".fa-pencil")))
    edit_button.click()

def edit_user_access_rights(browser):
    wait = WebDriverWait(browser, 10)
    edit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".fa-lock")))
    edit_button.click()

def actions_in_user_access_rights(browser):
    wait = WebDriverWait(browser, 10)
    edit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".fa-pencil")))
    edit_button.click()



# Iterate over the rows in the Excel file to add users
for row_index in range(df.shape[0]):
    role = df.loc[row_index, 'role']
    if role == 'client':
        ##  Call the functions for each row if role is client
        # click_add_user_button(browser)
        search_user(browser, df, row_index)
        edit_user(browser)
        # add_user_details(browser, 'users.xlsx', row_index)
        # select_timezone(browser, "(UTC+01:00) Brussels, Copenhagen, Madrid, Paris")
        # select_language(browser)
        # select_role(browser,role)
        # select_client(browser, 'users.xlsx', row_index)
        # additional_roles(browser, df, row_index)
        additional_roles_edit_user(browser, df, row_index)
        # click_activated_radio_button(browser)
        # click_create_user_button(browser)
        click_save_user_button(browser)
        # add_timestamp(df)
        # navigate_to_users_tab(browser)
        
    elif role == 'supplier':
        # Call the functions for each row if role is supplier
        click_add_user_button(browser)
        add_user_details(browser, 'users.xlsx', row_index)
        select_timezone(browser, "(UTC+01:00) Brussels, Copenhagen, Madrid, Paris")
        select_language(browser)
        select_role(browser,role)
        select_supplier(browser)
        # additional_roles(browser, df, row_index)
        click_activated_radio_button(browser)
        click_create_user_button(browser)
        # add_timestamp(df)
        # navigate_to_users_tab(browser)
        
    else:
        # Handle unknown role
        print('did not find users')
