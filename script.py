import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import gspread as gc
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = None
    if os.path.exists('api_key.json'):
        creds = Credentials.from_authorized_user_file('api_key.json')
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                r'C:\Users\EL11_gazy\Desktop\script\apikey.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('api_key.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def read_google_sheet(creds, spreadsheet_id, range_name):
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range=range_name).execute()
    values = result.get('values', [])
    return values

options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

a = 113

try:
    # Authenticate with Google Sheets API for both sheets
    creds = authenticate_google_sheets()

    # Read data from Sheet 1
    values_sheet1 = read_google_sheet(creds, "15h6PFrAgrwkMHaNwBCahdISTrZ63MYOFp4xt3PJCe-c", "sheet_name!A1:C")

    # Read data from Sheet 2
    values_sheet2 = read_google_sheet(creds, "15h6PFrAgrwkMHaNwBCahdISTrZ63MYOFp4xt3PJCe-c", "sheet_name !O1:R")

    # Iterate over each row in Sheet 1 starting from the second row
    for index, row_sheet1 in enumerate(values_sheet1[1727:1730], start=1):
        email, password, country = row_sheet1

        # Open the initial page
        driver.get("https://www.flowcv.com/")

        # Navigate to the /dashboard page
        driver.get("https://www.flowcv.com/dashboard")

        # Find and click on the login button
        wait = WebDriverWait(driver, 10)
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Login')]")))
        login_button.click()

        # Find the email input field and input the email
        email_input = driver.find_element(By.XPATH, "//input[@name='email']")
        email_input.clear()
        email_input.send_keys(email)

        # Find the password input field and input the password
        password_input = driver.find_element(By.XPATH, "//input[@name='password']")
        password_input.clear()
        password_input.send_keys(password)

        # Find and click on the login button
        login_button = driver.find_element(By.XPATH, "//button[@type='submit' and contains(text(), 'Login')]")
        login_button.click()

        # Wait for the login process to complete
        time.sleep(1)  # Increase the sleep time if needed

        # Now navigate to /app/resume/content
        driver.get("https://www.flowcv.com/app/resume/content")


        # Wait for the "Add Content" button to be clickable
        add_content_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add Content')]")))
        add_content_button.click()

        # Wait for the element to be clickable
        professional_experience_heading = wait.until(EC.element_to_be_clickable((By.XPATH, "//h4[contains(text(), 'Professional Experience')]")))
        professional_experience_heading.click()



        # Iterate over rows in Sheet 2 starting from where we left off
        for row_sheet2 in values_sheet2[a:]:
            job_title, role_description, job_title2, role_description2 = row_sheet2

            # Find the job title input field and input the job title
            job_title_input = driver.find_element(By.XPATH, "//input[@name='jobTitle']")
            job_title_input.clear()
            job_title_input.send_keys(job_title)

                        # Find the country input field and input the country
            country_input = driver.find_element(By.XPATH, "//input[@name='country']")
            country_input.clear()
            country_input.send_keys(country)

            # Find the input field for describing the role and achievements
            role_achievements_input = driver.find_element(By.XPATH, "//div[contains(@class, 'ql-editor') and contains(@class, 'ql-blank')]")
            role_achievements_input.clear()
            role_achievements_input.send_keys(role_description)

            # Wait for the parent div to be visible
            parent_div = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "flex.space-x-7")))

            # Find the save button within the parent div
            save_button = parent_div.find_element(By.XPATH, ".//button[@type='submit']")

            # Click the save button
            save_button.click()

            # Click "Add Content" again to add another job experience
            add_content_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add Content')]")))
            add_content_button.click()

            # Wait for the "Professional Experience" heading to be clickable again
            professional_experience_heading = wait.until(EC.element_to_be_clickable((By.XPATH, "//h4[contains(text(), 'Professional Experience')]")))
            professional_experience_heading.click()

            # Find the job title input field and input the second job title
            job_title_input = driver.find_element(By.XPATH, "//input[@name='jobTitle']")
            job_title_input.clear()
            job_title_input.send_keys(job_title2)


                        # Find the country input field and input the country
            country_input = driver.find_element(By.XPATH, "//input[@name='country']")
            country_input.clear()
            country_input.send_keys(country)

            # Find the input field for describing the second role and achievements
            role_achievements_input = driver.find_element(By.XPATH, "//div[contains(@class, 'ql-editor') and contains(@class, 'ql-blank')]")
            role_achievements_input.clear()
            role_achievements_input.send_keys(role_description2)

            # Wait for the parent div to be visible
            parent_div = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "flex.space-x-7")))

            # Find the save button within the parent div
            save_button = parent_div.find_element(By.XPATH, ".//button[@type='submit']")

            # Click the save button
            save_button.click()

            
            a += 1

            break  # Only process the next row of Sheet 2

        # After processing all job experiences for the current user, log out
        # Navigate back to the dashboard
        driver.get("https://www.flowcv.com/dashboard")

        # Click on the Logout button
        logout_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Logout')]")))
        logout_button.click()

        #a += 1

except Exception as e:
    print("Error:", e)
