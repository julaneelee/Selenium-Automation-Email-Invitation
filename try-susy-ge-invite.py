import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

###################
# 1 CONFIG
###################
URL_input = "https://susy.mdpi.com/special_issue/process/1760628"
EXCEL_FILE = "D:/Manuscripts/Special Issue work/GE invitation list/5.3.2026-GE.xlsx"
JOURNAL_TEXT = "Journal of Personalized Medicine (JPM)"
email_section_path='//*[@id="form_email"]'
NEXT_button_path='//*[@id="guestNextBtn"]'
back_button_path='//*[@id="specialBackBtn"]'
proceed_button_path='//*[@id="process-special-issue-guest-editor"]'
#next_special_issue="//a[@title='Next Special Issue']"
next_special_issue="/html/body/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[1]/a[2]/i"
###################
# 2. Load list of email 
# ################### 
try:
    df = pd.read_excel(EXCEL_FILE)
    # sort available time 
    df["Available Date (+90d6hrs)"] = pd.to_datetime(df["Available Date (+90d6hrs)"])
    df = df.sort_values("Available Date (+90d6hrs)")
    print()
except:
    print("Error in Excel")

###################
# 3 Setup Chrome - Logging in 
###################

driver = webdriver.Chrome()
driver.get(URL_input)
print("Logging in request")
wait = WebDriverWait(driver, 120)  # wait up to 5 minutes
wait.until(EC.presence_of_element_located((By.XPATH, email_section_path)))
print("Login detected. Continuing script...")

for index, row in df.iterrows():
    email = row["Email"]
    target_time = row["Available Date (+90d6hrs)"]

    while datetime.now() < target_time:
        time.sleep(0.1)

    print(f"\n>>> Processing: {email}")

    success = False
    while not success:
        try:
            # 1. Input Email
            email_input = wait.until(EC.element_to_be_clickable((By.XPATH, email_section_path)))
            email_input.clear()
            email_input.send_keys(email)
            driver.find_element(By.XPATH, NEXT_button_path).click()

            # 2. Wait for page to stabilize
            time.sleep(2) 
            
            # --- PRIORITY 1: Check for Limit Reached ---
            if "The number of proposed GE cannot exceed 5" in driver.page_source:
                print("Limit reached! Moving to next issue...")
                next_btn = wait.until(EC.element_to_be_clickable((By.XPATH, next_special_issue)))
                driver.execute_script("arguments[0].click();", next_btn)
                time.sleep(3) 
                continue # Retry same email on new page

            # --- PRIORITY 2: Check for PROCEED (The "Invite" case) ---
            # We check this FIRST to avoid accidental "Back" clicks
            proceed_btns = [btn for btn in driver.find_elements(By.XPATH, proceed_button_path) if btn.is_displayed()]
            if proceed_btns:
                print("Proceed button found! Inviting...")
                driver.execute_script("arguments[0].click();", proceed_btns[0])
                success = True
                time.sleep(2)
                continue

            # --- PRIORITY 3: Check for BACK (The "Already Invited" case) ---
            back_btns = [btn for btn in driver.find_elements(By.XPATH, back_button_path) if btn.is_displayed()]
            if back_btns:
                print("Already invited. Clicking Back.")
                driver.execute_script("arguments[0].click();", back_btns[0])
                success = True
                time.sleep(1)

        except Exception as e:
            print(f"Error: {str(e)[:50]}... Refreshing.")
            driver.refresh()
            time.sleep(2)