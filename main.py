from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from pprint import pprint
from time import sleep
import requests
import logging
import os
from dotenv import load_dotenv
from creds import (
    web_url, calling_url,
    contact_name, company_phone, city, company_email, calling_company_address, calling_company_name, calling_company_url, call_count, note, zipcode, service_provider
)
from fetch_email import fetch_otp, login, logout
from datetime import datetime
import pytz

# Load environment variables from .env file
load_dotenv()
g_api = os.getenv('GAPI')

# Create a folder for logging
log_folder = 'logs'
if not os.path.exists(log_folder):
    os.makedirs(log_folder)

# Create a subfolder with the current date
current_date = datetime.now().strftime('%Y-%m-%d')
subfolder = os.path.join(log_folder, current_date)
if not os.path.exists(subfolder):
    os.makedirs(subfolder)

# Set up logging
log_file = os.path.join(subfolder, 'automation.log')
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def driver_setup():
    """
    Set up a Chrome WebDriver with specified options for optimization.

    Parameters:
    None

    Returns:
    driver (webdriver.Chrome): The initialized Chrome WebDriver instance.

    The function initializes a Chrome WebDriver with various options for optimization.
    These options include disabling GPU hardware acceleration, disabling extensions,
    disabling popup blocking, disabling images, disabling JavaScript, disabling infobars,
    maximizing the window, and disabling software rasterizer.
    """
    # Set up Chrome options for optimization
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument('--headless')  # Run in headless mode
    chrome_options.add_argument('--disable-gpu')  # Disable GPU hardware acceleration
    chrome_options.add_argument('--no-sandbox')  # Bypass OS security model
    chrome_options.add_argument('--disable-dev-shm-usage')  # Overcome limited resource problems
    chrome_options.add_argument('--disable-extensions')  # Disable extensions
    chrome_options.add_argument('--disable-popup-blocking')  # Disable popup blocking
    chrome_options.add_argument('--disable-images')  # Disable images
    # chrome_options.add_argument('--disable-javascript')  # Disable JavaScript
    chrome_options.add_argument('--disable-infobars')  # Disable infobars
    chrome_options.add_argument('--start-maximized')  # Start maximized
    chrome_options.add_argument('--disable-software-rasterizer')  # Disable software rasterizer

    # Initialize the WebDriver
    driver = webdriver.Chrome(options=chrome_options)
    logging.info('Driver Initiated!')
    return driver

def get_sheet_data(id):
    """
    Fetch data from a Google Sheet using the provided ID.

    Parameters:
    id (str): The ID of the Google Sheet.

    Returns:
    dict: The JSON response from the Google Sheets API, or None if an error occurred.

    This function sends a GET request to the Google Sheets API using the provided ID and API key.
    It then checks the response status code and returns the JSON response if the status code is 200.
    If an error occurs during the request, it prints an error message and returns None.
    """
    try:
        response = requests.get(f'https://sheets.googleapis.com/v4/spreadsheets/{id}/values/!A1:Z?alt=json&key={g_api}')
        
        if response.status_code == 200:
            logging.info('Data Fetched From Google Sheet!')
            return response.json()
            
    except Exception as e:
        print(f'Error fetching data from Google Sheet: {str(e)}')
        
def fill_form(driver, phone_numbers, current_time, mail):
    """
    Fills the form on the website with the provided phone numbers, current time, and email.
    It also handles the OTP verification process and returns the feedback ID.

    Parameters:
    driver (webdriver.Chrome): The initialized Chrome WebDriver instance.
    phone_numbers (list): A list of phone numbers to be filled in the form.
    current_time (datetime): The current time in UTC.
    mail (str): The email address used for OTP verification.

    Returns:
    str: The feedback ID generated after submitting the form.
    """
    for i, phone_number in enumerate(phone_numbers):
        if i >= 20:
            break
        # Find the element and send the Enter key
        num_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, f'enterprise_phone_{i}')))
        sleep(0.2)
        num_input.send_keys(phone_number)
        if i < min(len(phone_numbers), 20) - 1:
            try:
                print(i < min(len(phone_numbers), 20) - 1)
                # Add Number Button
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'add-number-command'))).click()
            except Exception as e:
                print(f"Error occurred while clicking 'add-number-command': {e}")
                break
    
    try:
        # Select Category
        dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_category')))
        select = Select(dropdown)
        select.select_by_value('telemarketing')
        
        # Fill Other Details
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_contact_name'))).send_keys(contact_name)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_contact_phone'))).send_keys(company_phone)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_contact_email'))).send_keys(company_email)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_company_name'))).send_keys(calling_company_name)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_company_address_line_1'))).send_keys(calling_company_address)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_company_address_city'))).send_keys(city)
        
        # DropDown of State
        state_dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_company_address_state')))
        state_select = Select(state_dropdown)
        state_select.select_by_value('FL')
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_company_address_zip'))).send_keys(zipcode)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_company_url'))).send_keys(calling_company_url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'enterprise_service_provider'))).send_keys(service_provider)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'call_count'))).send_keys(call_count)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'additional_feedback'))).send_keys(note)
        
        send_otp_ = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'send-verification-code')))
        send_otp_.click()
        
        # Get OTP!
        for _ in range(3):
            print("Waiting For OTP...")
            sleep(5)
            otp = fetch_otp(mail, current_time)
            if otp:
                otp = otp.split(':')[-1].replace(' ', '')
                logging.info(f"OTP Received: {otp}")
                break
            else:
                print("No OTP Received Yet. Waiting For OTP...")
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'captcha'))).send_keys(otp)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'submitButton'))).click()
        # Get Feedback Number
        feedback_ = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, 'feedback-id'))).text.strip()
        return feedback_
    except Exception as e:
        logging.error(f"Error occurred while filling form: {e}")
        return None


def process_batches(values):
    """
    Divides a list of values into smaller batches of 20.

    Parameters:
    values (list): A list of values to be divided into batches.

    Returns:
    list: A list of lists, where each inner list represents a batch of 20 values.
          If the number of values is not a multiple of 20, the last batch may contain fewer than 20 values.
    """
    batches = [values[i:i + 20] for i in range(0, len(values), 20)]
    logging.info('Made Batches of 20 Numbers for Further Processing')
    return batches

def start_submission(driver):
    """
    This function navigates to the specified web page URL and clicks on the registration button.

    Parameters:
    driver (webdriver.Chrome): The initialized Chrome WebDriver instance.

    Returns:
    None

    The function performs the following steps:
    1. Navigates to the web page URL using the provided WebDriver instance.
    2. Refreshes the current page (Page Does not load on 1st Attempt, Restricted by URL).
    3. Waits for the registration button to be clickable using WebDriverWait and the specified locator.
    4. Clicks on the registration button.
    """
    # Go to the Page URL
    driver.get(web_url)
    driver.refresh()

    # Register Button
    reg_num_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'nextButton')))
    reg_num_button.click()


def main(driver, values):
    """
    This function orchestrates the entire process of logging in, fetching data from a Google Sheet,
    processing the data into batches, starting the submission process, filling the form, and logging out.

    Parameters:
    driver (webdriver.Chrome): The initialized Chrome WebDriver instance.
    values (list): A list of values fetched from the Google Sheet. Each value represents a batch of phone numbers.

    Returns:
    None

    The function performs the following steps:
    1. Logs in to the email account using the `login()` function.
    2. Initializes an empty list `results_list` to store the results.
    3. Processes the `values` into batches of 20 using the `process_batches()` function.
    4. Iterates over each batch and performs the following steps:
        - Calls the `start_submission()` function to navigate to the web page and click on the registration button.
        - Initializes an empty dictionary `result` to store the batch and feedback ID.
        - Retrieves the current time in UTC.
        - Extracts the phone numbers from the current batch.
        - Calls the `fill_form()` function to fill the form with the phone numbers, current time, and email.
        - Stores the feedback ID in the `result` dictionary.
        - Prints the `result` dictionary.
        - Waits for 1 second before proceeding with the next batch.
    5. Closes the WebDriver after the task is done using the `logout()` and `driver.quit()` functions.
    6. Prints a separator line.
    7. Prints the `results_list`.
    """
    try:
        mail = login()
        results_list = []
        
        # Process values in batches of 20
        batches = process_batches(values)
        for index, batch in enumerate(batches):
            print(f'Processing Batch {index + 1}/{len(batches)}')
            logging.info(f'Processing Batch {index + 1}/{len(batches)}')
            start_submission(driver)
            result = {}
            result['Batch'] = batch
            # Get the current time in UTC
            current_time = datetime.now(pytz.utc)
            print(current_time)
            phone_numbers = [row[0] for row in batch]
            feedback = fill_form(driver, phone_numbers, current_time, mail)
            if feedback:
                result['Feedback ID'] = feedback
                logging.info(f'Batch {index + 1}/{len(batches)} Finished Successfully')
                # Wait before proceeding with the next batch
                sleep(0.5)
            else:
                logging.error(f'Batch {index + 1}/{len(batches)} RETURNED WITH AN ERROR!!!')
                continue
    finally:
        # Close the WebDriver & Mail Connection after the task is done
        logout(mail)
        driver.quit()
        print('---------------------------------')
        pprint(results_list)

if __name__ == "__main__":
    '''
        Sheet ID is stored in Sheet URL
        The Pattern of Sheet URL is: https://docs.google.com/spreadsheets/d/{sheet_id}/edit?gid=0#gid=0
        e.g. https://docs.google.com/spreadsheets/d/1YyrMfGEl2KioO90rcKYuqFrzJhMblJsTmOyxvYhZtvo/edit?gid=0#gid=0
    '''
    
    sheet_id = "1YyrMfGEl2KioO90rcKYuqFrzJhMblJsTmOyxvYhZtvo"
    json_response = get_sheet_data(sheet_id)
    
    if json_response:
        # Extract the 'values' key from the response
        values = json_response.get('values', [])
        
        if values:
            # Remove the header row if present
            values = values[1:]
        
        driver = driver_setup()
        main(driver, values)
