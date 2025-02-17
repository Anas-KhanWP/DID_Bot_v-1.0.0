import email
import imaplib
import os
import logging
from dotenv import load_dotenv
from email.utils import parsedate_to_datetime
from datetime import datetime
import pytz

# Load environment variables from .env file
load_dotenv()

# Get credentials from environment variables
webmail_url = os.getenv('WEBMAIL_URL')
web_mail = os.getenv('WEB_MAIL')
password = os.getenv('PASSWORD')
server_ = os.getenv('SERVER')

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

def login():
    """
    Connects to the mail server using IMAP4_SSL protocol, logs in with provided credentials,
    and selects the inbox.

    Parameters:
    None

    Returns:
    mail (imaplib.IMAP4_SSL): An instance of the IMAP4_SSL class representing the connected mail server.
    """
    # Connect to the mail server
    mail = imaplib.IMAP4_SSL(server_)
    mail.login(web_mail, password)
    logging.info('Logged in to the mail server successfully.')
    return mail

def logout(mail):
    """
    Logs out from the mail server.

    Parameters:
    mail (imaplib.IMAP4_SSL): An instance of the IMAP4_SSL class representing the connected mail server.

    Returns:
    None
    """
    mail.logout()
    logging.info('Logged out from the mail server.')

def fetch_otp(mail, start_time):
    """
    Fetches the subject of the most recent email from a specific sender within a given time frame.

    Parameters:
    mail (imaplib.IMAP4_SSL): An instance of the IMAP4_SSL class representing the connected mail server.
    start_time (datetime.datetime): The start time of the time frame to search for emails.

    Returns:
    str: The subject of the most recent email from the specific sender within the given time frame.
         Returns None if no emails are found or an error occurs.
    """
    try:
        # Select the inbox
        mail.select('inbox')

        # Search for emails from the specific sender
        search_criteria = '(FROM "no-reply@tnsi.com")'
        status, data = mail.search(None, search_criteria)

        if status != 'OK':
            logging.error('Search failed.')
            return

        mail_ids = data[0].split()
        
        if not mail_ids:
            logging.info('No emails found from no-reply@tnsi.com.')
            return

        # Fetch headers and filter by date
        email_dates = []
        # for i in mail_ids:
        status, data = mail.fetch(mail_ids[-1], '(BODY[HEADER.FIELDS (DATE)])')
        # if status != 'OK':
        #     logging.error(f'Failed to fetch email ID {i}')
        #     continue
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                date = parsedate_to_datetime(msg['Date'])
                
                # print(f"Comparison| Start_time => {start_time} | Message_time => {date}")
                
                # Ensure email date is in UTC
                if date.tzinfo is None:
                    date = date.replace(tzinfo=pytz.utc)
                
                if date > start_time:
                    email_dates.append((mail_ids[-1], date))

        if not email_dates:
            logging.info('Still Waiting For OTP.')
            return

        # Sort emails by date, most recent first
        email_dates.sort(key=lambda x: x[1], reverse=True)
        most_recent_email_id = email_dates[0][0]

        # Fetch the most recent email by ID
        status, data = mail.fetch(most_recent_email_id, '(RFC822)')
        if status != 'OK':
            logging.error('Failed to fetch the most recent email.')
            return

        for response_part in data:
            if isinstance(response_part, tuple):
                message = email.message_from_bytes(response_part[1])
                mail_subject = message['subject']
                return mail_subject
    except Exception as e:
        logging.error(f'An error occurred while fetching email: {e}')
        print(f'An error occurred while fetching email: {e}')

if __name__ == '__main__':
    current_time = datetime.now(pytz.utc).replace(microsecond=0)
    mail = login()
    from time import sleep
    while True:
        sleep(5)
        otp = fetch_otp(mail, current_time)
        if otp:
            print(otp)
            break
        else:
            print('retrying')
        
    logout(mail)
