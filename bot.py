import tempfile
from seleniumbase import Driver
from selenium.common.exceptions import *
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import os, json, os.path, re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from plyer import notification
import pandas as pd
from openpyxl import load_workbook

SCOPES = ["https://mail.google.com/"]

driver = Driver(uc=True, headless=False)
cookie_json = 'cookies.json'

excel_file = 'passports.xlsx'
sheet_name = 'Sheet1'
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
for column in df.select_dtypes(include=['datetime64']).columns:
  df[column] = df[column].dt.strftime('%Y-%b-%d')
json_data = df.to_json(orient='records')
output_json_file = 'passport.json'

with open(output_json_file, 'w') as json_file:
  json_file.write(json_data)
print("Excel data successfully converted to JSON with dates in YYYY-MM-DD format.")
  
with open('profile.json', 'r') as file:
  profile = json.load(file)
  
with open('passport.json', 'r') as info:
  passport = json.load(info)
  
login_url = profile['login_url']
dashboard_url = profile['dashboard_url']

def get_password():
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
      token.write(creds.to_json())
  try:
    service = build("gmail", "v1", credentials=creds)
    sender_email = "donotreply@vfshelpline.com"
    recipient_email = profile['email']
    query = f"from:{sender_email} to:{recipient_email}"
    results = service.users().messages().list(userId="me", q=query).execute()
    messages = results.get("messages", [])
    if not messages:
      print("No messages found.")
      return ""
    print("Messages sent to the recipient:")
    first_message_id = messages[0]['id']
    msg = service.users().messages().get(userId='me', id=first_message_id).execute()
    snippet = msg['snippet']
    otp_match = re.search(r'\b\d{6}\b', snippet)
    if otp_match:
      otp = otp_match.group(0)
      print(f"Extracted OTP: {otp}")
      service.users().messages().delete(userId='me', id=first_message_id).execute()
      print("Message deleted successfully.")
      return otp
    else:
      print("OTP not found in the snippet.")
      return ""
  except HttpError as error:
    print(f"An error occurred: {error}")

def wait_url(url):
  while True:
    cur_url = driver.current_url
    if cur_url == url:
      break
    
def password_input(keyboard_rows, char):
  char_found = False
  for row in keyboard_rows:
    try:
        buttons = row.find_elements(By.CSS_SELECTOR, 'button[class^="touch-keyboard-key standard-key"]')
        for button in buttons:
            if button.text.strip() == char:
                button.click()
                print("clicked ->", char)
                char_found = True
                break
        if char_found:
            return char_found

    except StaleElementReferenceException:
        print("Stale elements detected, retrying...")
        break
      
def notify(title, message, time):
  notification.notify(
    title=title,
    message=message,
    app_name="Visa Booking",
    timeout=time
  )
  
def checkParagraph():
  try:
    paragraph = driver.find_element(By.CSS_SELECTOR, 'button[class^="position-relative"]')
    paragraph.click()
    paragraphText = driver.find_element(By.CSS_SELECTOR, 'p[class="fs-sm-17 c-brand-grey-para mb-30"]')
    txt = paragraphText.text.strip()
    if 'under process' in txt:
      return "UnderProcessing"
    elif 'limit restricted' in txt:
      return "Restrict"
    else:
      return "None"
  except NoSuchElementException:
    return "NoFound"

def updatePassportData(rowNum):
  excel_file_path = 'passports.xlsx'
  sheet_name = 'Sheet1'
  workbook = load_workbook(excel_file_path)
  worksheet = workbook[sheet_name]
  worksheet[f'O{rowNum}'] = "Yes"
  workbook.save(excel_file_path)
  print("Passport Excel file has been successfully updated.")
  
def login():
  try:
    driver.maximize_window()
    driver.get(login_url)
    if initialRound:
      cookies = WebDriverWait(driver, 10).until(
        lambda d: driver.get_cookies()
      )
      print("----------------------- cookies saved -----------------------")
      sleep(0.5)
      cookie_btn = WebDriverWait(driver, 10).until(
          EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[id="onetrust-reject-all-handler"]'))
      )
      cookie_btn.click()
      print("----------------------- cookie button clicked -----------------------")
      sleep(0.5)
    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='email']"))
    )
    email_input.send_keys(profile['email'])
    print("----------------------- email inputed -----------------------")
    sleep(0.5)
    password_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
    )
    if 'ago' not in login_url and 'prt' in login_url:
      password_element.send_keys(profile['password'])
    elif 'ago' in login_url and 'prt' in login_url:
      password_element.click()
      sleep(0.5)
      for char in profile['password']:
        char_found = False
        try:
          if char.isupper() or char.islower():
            while not char_found:
              WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="touch-keyboard-row ng-star-inserted"]'))
              )
              keyboard_rows = driver.find_elements(By.CSS_SELECTOR, 'div[class="touch-keyboard-row ng-star-inserted"]')
              char_found = password_input(keyboard_rows, char)
              if not char_found:
                latest_div = keyboard_rows[-1]
                buttons = latest_div.find_elements(By.TAG_NAME, 'button')
                buttons[0].click()
          else:
            while not char_found:
              WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="touch-keyboard-row ng-star-inserted"]'))
              )
              keyboard_rows = driver.find_elements(By.CSS_SELECTOR, 'div[class="touch-keyboard-row ng-star-inserted"]')
              char_found = password_input(keyboard_rows, char)
              if not char_found:
                latest_div = keyboard_rows[-1]
                buttons = latest_div.find_elements(By.TAG_NAME, 'button')
                buttons[1].click()
        except TimeoutException:
          print("Keyboard element not found.")
    print("----------------------- password inputed -----------------------")
    sleep(0.5)
    email_input.click()
    sleep(0.5)
    driver.uc_gui_click_captcha()
    sleep(5)
    driver.uc_gui_click_captcha()
    print("----------------------- passed captcha -----------------------")
    sleep(2)
    login_btn_clickable = False
    while not login_btn_clickable:
      try:
        login_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.mat-mdc-button-base'))
        )
        login_btn_clickable = True
      except TimeoutException:
        print("Still waiting for the button to be clickable...")
    login_btn.click()
    print("----------------------- login button clicked -----------------------")
    if 'ago' in login_url and 'prt' in login_url:
      print("---------------- waiting confirm password ----------------")
      sleep(10)
      confirm_password = ""
      while confirm_password == "":
        confirm_password = get_password()
      print("confirm_password -> ", confirm_password)
      print("--------------------- received confirm password ---------------------")
      confirm_password_element = WebDriverWait(driver, 10).until(
          EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
      )
      confirm_password_element.click()
      sleep(0.5)
      for char in confirm_password:
        num_found = False
        try:
          while not num_found:
            WebDriverWait(driver, 20).until(
              EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="touch-keyboard-row ng-star-inserted"]'))
            )
            keyboard_rows = driver.find_elements(By.CSS_SELECTOR, 'div[class="touch-keyboard-row ng-star-inserted"]')
            num_found = password_input(keyboard_rows, char)
        except TimeoutException:
          print("Keyboard element not found.")
          # return
      enter_btn = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="c-brand-grey-para fs-sm-16 mb-20 ng-star-inserted"]'))
      )
      enter_btn.click()
      print("--------------------- inputed confirm password ---------------------")
      sleep(0.5)
      driver.uc_gui_click_captcha()
      sleep(5)
      driver.uc_gui_click_captcha()
      print("------------------------ passed captcha ------------------------")
      sleep(2)
      signin_btn_clickable = False
      while not signin_btn_clickable:
        try:
          signin_btn = WebDriverWait(driver, 20).until(
              EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.mat-mdc-button-base'))
          )
          signin_btn_clickable = True
        except TimeoutException:
          print("Still waiting for the button to be clickable...")

      signin_btn.click()
      print("------------------------- signin clicked -------------------------")
    sleep(5)
  except Exception as e:
    print(f"An error occurred during login: {e}")
  finally:
    should_quit = False
    if should_quit:
      print("Quitting driver ..........")
      driver.quit()
    else:
      pass
    
def appointmentDetails():
  try:
    startbooking_btn = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, 'button[class^="btn custom-height-button mat-btn-lg btn-brand-orange"]'))
    )
    startbooking_btn.click()
    print("---------------------- start booking clicked ----------------------")
    sleep(5)
    if 'ago' in login_url and 'prt' in login_url:
      chooseAppointmentCategory_drop = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="mat-mdc-select-arrow-wrapper ng-tns-c91-9"]'))
      )
    elif 'ago' not in login_url and 'prt' in login_url:
      chooseAppointmentCategory_drop = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="mat-mdc-select-arrow-wrapper ng-tns-c91-8"]'))
      )
    chooseAppointmentCategory_drop.click()
    print("---------------------- chooseAppointmentCategory_drop ----------------------")
    if 'ago' in login_url and 'prt' in login_url:
      chooseAppointmentCategory_selects = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class^="ng-trigger ng-trigger-transformPanel ng-tns-c91-9"]'))
      )
      chooseAppointmentCategory_selects_options = chooseAppointmentCategory_selects.find_elements(By.TAG_NAME, 'mat-option')
      if profile['AppCategory'] == 'national':
        chooseAppointmentCategory_selects_options[0].click()
      elif profile['AppCategory'] == 'sengen':
        chooseAppointmentCategory_selects_options[1].click()
    elif 'ago' not in login_url and 'prt' in login_url:
      chooseAppointmentCategory_selects = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class^="ng-trigger ng-trigger-transformPanel ng-tns-c91-8"]'))
      )
      chooseAppointmentCategory_selects_options = chooseAppointmentCategory_selects.find_elements(By.TAG_NAME, 'mat-option')
      chooseAppointmentCategory_selects_options[1].click()
    print("---------------------- chooseAppointmentCategory_select[1] ----------------------")
    sleep(3)
    if 'ago' in login_url and 'prt' in login_url:
      chooseSubCategory_drop =  WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="mat-mdc-select-arrow-wrapper ng-tns-c91-7"]'))
      )
    elif 'ago' not in login_url and 'prt' in login_url:
      chooseSubCategory_drop =  WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="mat-mdc-select-arrow-wrapper ng-tns-c91-6"]'))
      )
    chooseSubCategory_drop.click()
    print("---------------------- chooseSubCategory_drop ----------------------")
    if 'ago' in login_url and 'prt' in login_url:
      chooseSubCategory_select = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class^="ng-trigger ng-trigger-transformPanel ng-tns-c91-7"]'))
      )
      chooseSubCategory_select.click()
    elif 'ago' not in login_url and 'prt' in login_url:
      chooseSubCategory_selects = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class^="ng-trigger ng-trigger-transformPanel ng-tns-c91-6"]'))
      )
      chooseSubCategory_selects_options = chooseSubCategory_selects.find_elements(By.TAG_NAME, 'mat-option')
      chooseSubCategory_selects_options[0].click()
    print("---------------------- chooseSubCategory_select ----------------------")
    continue1_btn_clickable = False
    while not continue1_btn_clickable:
      try:
        continue1_btn = WebDriverWait(driver, 20).until(
          EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-block btn-brand-orange mdc-button"]'))
        )
        continue1_btn_clickable = True
      except TimeoutException:
        print("Still waiting for the button to be clickable...")
    continue1_btn.click()
    print("---------------------- continue1 button clicked ----------------------")
  except Exception as e:
    print(f"An error occurred during appointmentDetails input: {e}")
  finally:
    should_quit = False
    if should_quit:
      print("Quitting driver ..........")
      driver.quit()
    else:
      pass
def yourDetails(passport0, applicantNum, index, jumped):
  print(f"--- applicant Number -> {applicantNum} ---")
  print(f"--- index -> {index} ---")
  try:
    if not jumped:
      if applicantNum == 0:
        sleep(29)
      else:
        sleep(25)
    firstName = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Enter your first name']"))
    )
    firstName.clear()
    firstName.send_keys(passport0['FirstName'])
    lastName = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Please enter last name.']"))
    )
    lastName.clear()
    lastName.send_keys(passport0['LastName'])
    gender = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, f"div[id='mat-select-value-{7+applicantNum*4}']"))
    )
    gender.click()
    gender_option = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, f"div[id='mat-select-{6+applicantNum*4}-panel']"))
    )
    gender_options = gender_option.find_elements(By.TAG_NAME, 'mat-option')
    if passport0['Gender'] == 'Female':
      gender_options[0].click()
    elif passport0['Gender'] == 'Male':
      gender_options[1].click()
    else:
      gender_options[2].click()
    birth_date = passport0["Birth"]
    year, month, day = birth_date.split('-')
    day = day.lstrip('0')
    birthday = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='dateOfBirth']"))
    )
    birthday.click()
    birthday_year = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='Select year']"))
    )
    birthday_year.click()
    birthday_year_options = birthday_year.find_elements(By.TAG_NAME, 'option')
    for birthday_year_option in birthday_year_options:
      if birthday_year_option.text.strip() == year:
        birthday_year_option.click()
        break
    birthday_month = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='Select month']"))
    )
    birthday_month.click()
    birthday_month_options = birthday_month.find_elements(By.TAG_NAME, 'option')
    for birthday_month_option in birthday_month_options:
      if birthday_month_option.text.strip() == month:
        birthday_month_option.click()
        break
    birthday_date_table = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='ngb-dp-month ng-star-inserted']"))
    )
    birthday_date_table_weeks = birthday_date_table.find_elements(By.CSS_SELECTOR, "div[class='ngb-dp-week ng-star-inserted']")
    birth_date_clicked = False
    for birthday_date_table_week in birthday_date_table_weeks:
      birthday_date_table_dates = birthday_date_table_week.find_elements(By.CSS_SELECTOR, "div[class='btn-light ng-star-inserted']")
      for birthday_date_table_date in birthday_date_table_dates:
        if birthday_date_table_date.text.strip() == day:
          birthday_date_table_date.click()
          birth_date_clicked = True
          break
      if birth_date_clicked:
        break
    national = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, f"div[id='mat-select-value-{9+applicantNum*4}']"))
    )
    national.click()
    sleep(0.5)
    national_list = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, f"div[id='mat-select-{8+applicantNum*4}-panel']"))
    )
    nationals = national_list.find_elements(By.TAG_NAME, 'span')
    for nation in nationals:
      if nation.text.strip() == passport0['Nationality']:
        nation.click()
        break
    passportNumber = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Enter passport number']"))
    )
    passportNumber.clear()
    passportNumber.send_keys(passport0['Passport'])
    
    passport_date = passport0["PassportExp"]
    passport_year, passport_month, passport_day = passport_date.split('-')
    passport_day = passport_day.lstrip('0')
    passoprtExp = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='passportExpirtyDate']"))
    )
    passoprtExp.click()
    passoprtExp_year = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='Select year']"))
    )
    passoprtExp_year.click()
    passoprtExp_year_options = passoprtExp_year.find_elements(By.TAG_NAME, 'option')
    for passoprtExp_year_option in passoprtExp_year_options:
      if passoprtExp_year_option.text.strip() == passport_year:
        passoprtExp_year_option.click()
        break
    passoprtExp_month = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='Select month']"))
    )
    passoprtExp_month.click()
    passoprtExp_month_options = passoprtExp_month.find_elements(By.TAG_NAME, 'option')
    for passoprtExp_month_option in passoprtExp_month_options:
      if passoprtExp_month_option.text.strip() == passport_month:
        passoprtExp_month_option.click()
        break
    passoprtExp_date_table = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='ngb-dp-month ng-star-inserted']"))
    )
    passoprtExp_date_table_weeks = passoprtExp_date_table.find_elements(By.CSS_SELECTOR, "div[class='ngb-dp-week ng-star-inserted']")
    exp_date_clicked = False
    for passoprtExp_date_table_week in passoprtExp_date_table_weeks:
      passoprtExp_date_table_dates = passoprtExp_date_table_week.find_elements(By.CSS_SELECTOR, "div[class='btn-light ng-star-inserted']")
      for passoprtExp_date_table_date in passoprtExp_date_table_dates:
        if passoprtExp_date_table_date.text.strip() == passport_day:
          passoprtExp_date_table_date.click()
          exp_date_clicked = True
          break
      if exp_date_clicked:
        break
    contactNumPri = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='44']"))
    )
    contactNumPri.clear()
    contactNumPri.send_keys(passport0['ContactNumber_pri'])
    contactNumSec = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='012345648382']"))
    )
    contactNumSec.clear()
    contactNumSec.send_keys(passport0['ContactNumber_sec'])
    email = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Enter Email Address']"))
    )
    email.clear()
    email.send_keys(passport0['Email'])
    save_btn_clickable = False
    while not save_btn_clickable:
      try:
        save_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn-brand-orange'))
        )
        save_btn_clickable = True
      except TimeoutException:
        print("Still waiting for the save button to be clickable...")
    save_btn.click()
    sleep(5)
    state = checkParagraph()
    print("already booked person alert modal state -> ", state)
    if state == 'UnderProcessing':
      notify('Booking State', 'Current client is under processing\nJump to the next client', 10)
      passport[index]['BookingState'] = "Yes"
      with open('passport.json', 'w') as updating:
        json.dump(passport, updating, indent=4)
      updatePassportData(index+2)
      return "JumpNextPerson"
    elif state == 'Restrict':
      notify('Email Restrict Alert', 'Your email ID is restricted.\nTry again after a few hours\nPlease input "y" to quit bot', 30)
      exit_txt = input('Your email ID is restricted.\nTry again after a few hours\nPlease input "y" to quit bot\n')
      if exit_txt == 'y':
        print("Quitting driver ..........")
        return "Quit"
    anotherApp_btn_clickable = False
    while not anotherApp_btn_clickable:
      try:
        anotherApp_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-brand-orange"]'))
        )
        anotherApp_btn_clickable = True
      except TimeoutException:
        print("continue...")
        continue2_btn = WebDriverWait(driver, 20).until(
          EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-block btn-brand-orange"]'))
        )
        continue2_btn.click()
        return "NextStep"
    if applicantNum == profile['applicantNumber']-1:
      print("continue...")
      continue2_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-block btn-brand-orange"]'))
      )
      continue2_btn.click()
      return "NextStep"
    else :
      anotherApp_btn.click()
    return "AnotherPerson"
  except Exception as e:
    print(f"An error occurred during yourDetails input: {e}")
  finally:
    should_quit = False
    if should_quit:
      print("Quitting driver ..........")
      driver.quit()
    else:
      pass
def rebook():
  back_btn = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-block btn-outline-brand-orange"]'))
  )
  back_btn.click()
  sleep(5)
  print("rebooking...")
  rebookContinue_btn = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-block btn-brand-orange"]'))
  )
  rebookContinue_btn.click()
  bookAppointment()
def bookAppointment():
  try:
    print("book appointment started...")
    sleep(7)
    calender_month = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, 'tbody[role="presentation"]'))
    )
    find_available_date = False
    nextCounter = 0
    while not find_available_date:
      calender_weeks = calender_month.find_elements(By.CSS_SELECTOR, 'tr[role="row"]')
      for calender_week in calender_weeks:
        calender_days = calender_week.find_elements(By.CSS_SELECTOR, 'td.fc-day-future.fc-daygrid-day.date-availiable')
        if calender_days:
          calender_day = calender_days[0]
          calender_day_text = calender_day.find_element(By.CSS_SELECTOR, 'a[class="fc-daygrid-day-number"]')
          print(calender_day_text.text.strip())
          calender_day.click()
          find_available_date = True
          break
      if not find_available_date:
        print('There is no available date..............')
        next_month_btn = WebDriverWait(driver, 20).until(
          EC.presence_of_element_located((By.CSS_SELECTOR, 'button[title="Next month"]'))
        )
        if nextCounter == 2:
          rebook()
        else:
          next_month_btn.click()
          print('Next Month Button Clikced...')
          nextCounter += 1
          sleep(5)
    sleep(3)
    chooseAppointmentTimeTable = WebDriverWait(driver, 20).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, 'table[class="table ba-slot-table w-auto"]'))
    )
    chooseAppointmentTimeAvailables = chooseAppointmentTimeTable.find_elements(By.TAG_NAME, 'input')
    chooseAppointmentTimeAvailables[0].click()
    continueBookAppointment_btn = WebDriverWait(driver, 20).until(
      EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn custom-height-button mat-btn-lg"]'))
    )
    continueBookAppointment_btn.click()
    sleep(5)
  except Exception as e:
    print(f"An error occurred during yourDetails input: {e}")
  finally:
    should_quit = False
    if should_quit:
      print("Quitting driver ..........")
      driver.quit()
    else:
      pass
def service():
  try:
    continueService_btn = WebDriverWait(driver, 20).until(
      EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn custom-height-button mat-btn-lg"]'))
    )
    continueService_btn.click()
    sleep(5)
  except Exception as e:
    print(f"An error occurred during yourDetails input: {e}")
  finally:
    should_quit = False
    if should_quit:
      print("Quitting driver ..........")
      driver.quit()
    else:
      pass
def review():
  try:
    checkboxes = driver.find_elements(By.CSS_SELECTOR, 'input[type="checkbox"]')
    checkboxes[0].click()
    sleep(0.5)
    checkboxes[1].click()
    sleep(0.5)
    payOnline_btn = WebDriverWait(driver, 20).until(
      EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-block btn-brand-orange"]'))
    )
    payOnline_btn.click()
    sleep(1)
    continueReview_btns = driver.find_elements(By.CSS_SELECTOR, 'button[class^="btn mat-btn-lg btn-block btn-outline-brand-orange"]')
    continueReview_btns[1].click()
    sleep(0.5)
  except Exception as e:
    print(f"An error occurred during yourDetails input: {e}")
  finally:
    should_quit = False
    if should_quit:
      print("Quitting driver ..........")
      driver.quit()
    else:
      pass
def create_temp_profile():
  temp_dir = tempfile.mkdtemp()
  return temp_dir
def initialize_driver():
  temp_profile = create_temp_profile()
  driver = Driver(uc=True, headless=False, user_data_dir=temp_profile)
  return driver

if __name__ == "__main__":
  driver = initialize_driver()
  print("---------- cache and cookies deleted ----------")
  try:
    initialRound = True
    quitBooking = False
    while True:
      if os.path.exists(cookie_json):
        driver.get(dashboard_url)
        sleep(2)
        with open('cookies.json', 'r') as f:
          cookies = json.load(f)
        for cookie in cookies:
          cookie_dict = {
            'name': cookie['name'],
            'value': cookie['value'],
            'domain': cookie['domain'],
            'path': cookie['path'],
            'secure': cookie.get('secure', False),
            'httpOnly': cookie.get('httpOnly', False)
          }
          driver.add_cookie(cookie_dict)
      else:
        login()
        appointmentDetails()
        index = -1
        applicantNum = 0
        jumped = False
        for passport0 in passport:
          if passport0['AppCategory'] == profile['AppCategory']:
            index += 1
            if passport0['BookingState'] == 'No':
              detailedInfoState = yourDetails(passport0, applicantNum, index, jumped)
            else:
              detailedInfoState = "NextPerson"
            if detailedInfoState == "NextStep" or detailedInfoState == "Quit":
              break
            elif detailedInfoState == "JumpNextPerson":
              jumped = True
            elif detailedInfoState == "AnotherPerson" and passport0['BookingState'] == 'No':
              jumped = False
              applicantNum += 1
        if detailedInfoState == 'Quit':
          quitBooking = True
        if quitBooking == False:
          bookAppointment()
          service()
          review()
          sleep(10)
          payment_url = driver.current_url
          if 'token' in payment_url:
            notify('Booking Alert!', f'{passport[index]['No']}] {passport[index]['FirstName']}\'s Booking was successful', 10)
            print(f'{passport[index]['No']}] {passport[index]['FirstName']}\'s Booking was successful')
            notify('Payment Alert', f'{payment_url}', 30)
            passport[index]['BookingState'] = "Yes"
            with open('passport.json', 'w') as updating:
              json.dump(passport, updating, indent=4)
            updatePassportData(index+2)
          else:
            notify('Booking Alert!', 'Booking failed', 10)
            print(f'{passport[index]['No']}] {passport[index]['FirstName']}\'s Booking failed')
          print(payment_url)
          while True:
            response = input("Are you ready to continue?(y/n)\n")
            if response == 'y':
              initialRound = False
              break
            elif response == 'n':
              quitBooking = True
              break
      if quitBooking:
        break
  except Exception as general_exception:
    print(f"General exception caught: {general_exception}")
  finally:
    pass