#!/bin/bash python3
# -*- encoding: utf-8 -*-
# upload_ids.py


# Importing Python Built-In Modules
import email
import imaplib
import logging
import os
import os.path
import re
import time

# Importing Exterior Modules
import easygui
from selenium.webdriver.chrome.webdriver import WebDriver
import undetected_chromedriver as uc
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv


def wait_until_clickable(webdriver, xpath:str):
    return WebDriverWait(webdriver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath)))


def main(uspto_email:str, uspto_password:str, imap_server:str, imap_email:str, imap_password:str, sponsor:str):
    '''Main function for the upload_ids.py module.'''

    HOME = os.path.expanduser('~')

    # Initial Questions GUI
    while True:
        vals = easygui.multenterbox('Complete the following information, then hit OK.', 'Upload IDS Files', ['Application No.', 'Confirmation No.'])
        logging.info(vals)
        if vals is None:
            return
        nvals = []
        for val in vals:
            nvals.append(val.lower().strip())
        if len(nvals[0]) != 8:
            easygui.msgbox('Your Application Number was invalid: not 8 digits.', 'Upload IDS Files')
            continue
        if len(nvals[1]) != 4:
            easygui.msgbox('Your Confirmation Number was incorrect: not 4 digits.', 'Upload IDS FIles')
            continue
        else:
            goodvals = nvals
            break

    dirfile = easygui.diropenbox('Choose a directory with your IDS files', 'Upload IDS Files', HOME)
    idsfee = easygui.buttonbox('Are there IDS fees?', 'Upload IDS Files', ('Yes', 'No'), default_choice='No')

    applno = goodvals[0]
    confno = goodvals[1]

    # Time to wait
    x = 1.2

    # Opening Chrome
    logging.info('Launching Chromedriver')

    options = Options()
    options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:95.0) Gecko/20100101 Firefox/95.0')
    options.add_argument('--start-maximized')
    options.add_argument(f'user-data-dir={HOME}\\AppData\\Local\\Google\\Chrome\\User Data')

    driver = uc.Chrome(ChromeDriverManager().install(), options=options)

    logging.info('GET EFS Website')
    driver.get("https://efs-my.uspto.gov/EFSWebUIRegistered/EFSWebRegistered")

    # Email Field
    logging.info('CLICK/SENDKEYS Email Address Field')
    button = wait_until_clickable(driver, '//*[@id="siw-username"]')
    button.click()
    button.send_keys(uspto_email)
    button.send_keys(Keys.ENTER)
    # Password Field
    logging.info('CLICK/SENDKEYS Password Field')
    button = wait_until_clickable(driver, '//*[@id="siw-password"]')
    button.click()
    button.send_keys(uspto_password)
    # reCaptcha
    easygui.msgbox("Complete the reCAPTCHA; When done (or if there is no CAPTCHA), click the OK button to continue.", 'Upload IDS Files')

    try:
        button = wait_until_clickable(driver, '//*[@id="siw-enter-pwd-form"]/fieldset/div[4]/button')
        button.click()
    except NoSuchElementException:
        logging.warning('Skipping CLICK of next page because already clicked.')

    # 2-Step Authentication - 1st page
    logging.info('CLICK 2-Step Authentication - 1st page')
    button = wait_until_clickable(driver, '//*[@id="tfa-form"]/div/button')
    button.click()
    # 2-Step Authentication - Email
    # Logging into Email
    logging.info('IMAPLIB check email for 2-step verification code')
    M = imaplib.IMAP4_SSL(imap_server)
    user = imap_email
    password = imap_password
    M.login(user, password)
    M.select("INBOX")
    # Looking for Code Email
    logging.info('WHILE search email for Authentication code.')
    time.sleep(10)

    while True:
        time.sleep(2)
        year, month, day, _, _ = map(int, time.strftime("%Y %m %d %H %M").split())
        months = {'1':'Jan','2':'Feb','3':'Mar','4':'Apr','5':'May','6':'Jun',
        '7':'Jul','8':'Aug','9':'Sep','10':'Oct','11':'Nov','12':'Dec'}
        d = '-'
        if len(str(day)) == 1:
            day = '0' + str(day)

        _,data = M.search(None,('SUBJECT "Your authentication code"'),f'(ON "{str(day) + d + str(months[str(month)]) + d + str(year)}")','(UNSEEN)')

        try:
            _,email_data = M.fetch(data[0],"(RFC822)")

            raw_email = email_data[0][1]
            raw_email_string = raw_email.decode('UTF-8')
            email_message = email.message_from_string(raw_email_string)

            for part in email_message.walk():
                if part.get_content_type() == 'text/plain':
                    body = part.get_payload(decode=True)
                    body = body.decode("UTF-8")
                    pattern = r'Your authentication code is (\d{6})'
                    athun_code = re.search(pattern,body)
        except:
            logging.info('No email; retrying')
            continue
        else:
            # Closing IMAP
            M.close()
            M.logout()
            break

    # 2-Step Authentication 2nd Page
    logging.info('CLICK/SENDKEYS 2-Step Authentication 2nd Page')
    button = wait_until_clickable(driver, '//*[@id="tfa-pin"]')
    button.send_keys(athun_code.group(1))
    button = wait_until_clickable(driver, '//*[@id="tfa-pin-form"]/div/button[2]')
    button.click()

    # Logged in
    # Clicking E-File
    logging.info('CLICK E-File')
    button = wait_until_clickable(driver, '//*[@id="PatenteBusiness0"]/a')
    button.click()
    # Clicking E-File Registered
    button = wait_until_clickable(driver, '//*[@id="PatenteBusiness0"]/div/a[2]')
    button.click()
    # Swiching Tabs
    window = driver.window_handles[1]
    driver.switch_to.window(window)
    # Clicking Sponsor
    time.sleep(7)
    button = driver.find_element_by_link_text(sponsor)
    button.click()

    # Existing Application
    logging.info('Entering Application Stage')
    # Scroll Down
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    button = wait_until_clickable(driver, '//*[@id="Existing Application"]')
    # Documents/Fees
    button = wait_until_clickable(driver, '//*[@id="documents fees"]')
    button.click()
    # Application Number
    button = wait_until_clickable(driver, '//*[@id="appnumfollowon"]')
    button.click()
    button.send_keys(applno)
    # Confirmation Number
    button = wait_until_clickable(driver, '//*[@id="connumfollowon"]')
    button.click()
    button.send_keys(confno)
    # Continue
    button = wait_until_clickable(driver, '//*[@id="Submit"]')
    button.click()

    # Choosing Files:
    logging.info('Uploading Files')
    ind = 0
    # Foreign References
    dirfile = dirfile + '\\foreign'
    are = os.path.isdir(dirfile)
    if not are:
        pass
    else:
        # Looping through all files in the directory, getting their filenames, and filing them...
        for filename in os.listdir(dirfile):
            if filename.endswith(".pdf"):
                # Scroll Down
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
                # Inputting files
                button = wait_until_clickable(driver, f'//*[@id="upld{ind}"]')
                button.send_keys(os.path.join(dirfile, filename))
                # Selecting Category IDS/References
                select = Select(wait_until_clickable(driver, f'//*[@id="Category_{ind}"]'))
                select.select_by_value('IDS/References')
                # Selecting Foreign Refrences
                select = Select(wait_until_clickable(driver, f'//*[@id="DocDescription_{ind}"]'))
                select.select_by_value('Foreign Reference')
                # Adding Files
                button = wait_until_clickable(driver, '//*[@id="moredocument"]/tbody/tr/td[6]/input')
                button.click()
                time.sleep(x)
                ind += 1
            else:
                continue


    # Other Reference-Patent/App/Search documents
    inp = dirfile + '\\search'
    are = os.path.isdir(inp)
    if not are:
        pass
    else:
        # Looping through all files in the directory, getting their filenames, and filing them...
        for filename in os.listdir(inp):
            if filename.endswith(".pdf"):
                # Scroll Down
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
                # Inputting files
                button = wait_until_clickable(driver, f'//*[@id="upld{ind}"]')
                button.send_keys(os.path.join(inp, filename))
                # Selecting Category IDS/References
                select = Select(wait_until_clickable(driver, f'//*[@id="Category_{ind}"]'))
                select.select_by_value('IDS/References')
                # Selecting Other Reference-Patent/App/Search documents
                select = Select(wait_until_clickable(driver, f'//*[@id="DocDescription_{ind}"]'))
                select.select_by_value('Other Reference-Patent/App/Search documents')
                # Adding Files
                button = wait_until_clickable(driver, '//*[@id="moredocument"]/tbody/tr/td[6]/input')
                button.click()
                time.sleep(x)
                ind += 1
            else:
                continue


    # Non-Patent Literature
    inp = dirfile + '\\non'
    are = os.path.isdir(inp)
    if not are:
        pass
    else:
        # Looping through all files in the directory, getting their filenames, and filing them...
        for filename in os.listdir(inp):
            if filename.endswith(".pdf"):
                # Scroll Down
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
                # Inputting files
                button = wait_until_clickable(driver, f'//*[@id="upld{ind}"]')
                button.send_keys(os.path.join(inp, filename))
                # Selecting Category IDS/References
                select = Select(wait_until_clickable(driver, f'//*[@id="Category_{ind}"]'))
                select.select_by_value('IDS/References')
                # Selecting Non-Patent Literature
                select = Select(wait_until_clickable(driver, f'//*[@id="DocDescription_{ind}"]'))
                select.select_by_value('Non Patent Literature')
                # Adding Files
                button = wait_until_clickable(driver, '//*[@id="moredocument"]/tbody/tr/td[6]/input')
                button.click()
                time.sleep(x)
                ind += 1
            else:
                continue


    # IDS Form
    for filename in os.listdir(dirfile):
        if filename.endswith('.pdf'):
            inp = filename
        else:
            continue
    # Scroll Down
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")

    button = wait_until_clickable(driver, f'//*[@id="upld{ind}"]')
    button.send_keys(dirfile + "\\" + inp)
    # Selecting Category IDS/References
    select = Select(wait_until_clickable(driver, f'//*[@id="Category_{ind}"]'))
    select.select_by_value('IDS/References')
    # Selecting IDS Form
    select = Select(wait_until_clickable(driver, f'//*[@id="DocDescription_{ind}"]'))
    select.select_by_value('Information Disclosure Statement (IDS) Form (SB08)')
    # Adding Files
    button = wait_until_clickable(driver, '//*[@id="moredocument"]/tbody/tr/td[6]/input')
    button.click()

    # Upload & Validate
    button = wait_until_clickable(driver, '//*[@id="Submit"]')
    button.click()

    # Scroll Down
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    # Continue
    button = wait_until_clickable(driver, '//*[@id="Submit"]')
    button.click()

    # IDS Fees
    if idsfee.lower() == 'y':
        # Scroll Down
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
        button = wait_until_clickable(driver, '//*[@id="miscfeesbox"]')
        button.click()
        button = wait_until_clickable(driver, '//*[@id="miscfeesbox"]/option[4]')
        button.click()
        button = wait_until_clickable(driver, '//*[@id="miscfees"]/table[1]/tbody/tr/td[2]/div/a[1]/img')
        button.click()
    else:
        pass

    button = wait_until_clickable(driver, '//*[@id="Submit"]')
    button.click()
    # Saving for Later Submission
    button = wait_until_clickable(driver, '//*[@id="Save for Later Submission"]')
    button.click()
    # Alert Confirmation
    obj = driver.switch_to.alert
    obj.accept()

    logging.info('Complete; Reported error code 0')


if __name__ == '__main__':
    # Setup Logging
    i = 0
    while True:
        if os.path.isfile(f'upload_ids.py.log.{str(i)}'):
            i += 1
            continue
        else:
            FNAME = f'upload_ids.py.log.{str(i)}'
            with open(FNAME, 'a', encoding='utf-8') as a:
                pass
            break
    logging.basicConfig(filename=FNAME, filemode='a', format='%(asctime)s - %(message)s', level=logging.INFO)

    load_dotenv()
    USPTO_EMAIL = os.getenv('USPTO_EMAIL')
    USPTO_PASSWORD = os.getenv('USPTO_PASSWORD')
    IMAP_SERVER = os.getenv('IMAP_SERVER')
    IMAP_EMAIL = os.getenv('IMAP_EMAIL')
    IMAP_PASSWORD = os.getenv('IMAP_PASSWORD')
    SPONSOR = os.getenv('SPONSOR')

    try:
        main(USPTO_EMAIL, USPTO_PASSWORD, IMAP_SERVER, IMAP_EMAIL, IMAP_PASSWORD, SPONSOR)
    except Exception as e:
        logging.exception('')
        easygui.msgbox('''The program probably crashed.
Check your Desktop for the log files to see what happened.
If the webdriver is still running, by clicking "Ok", you will close the webdriver. You can finish the operation if you'd like, or troubleshoot and restart the program.''')
