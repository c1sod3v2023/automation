import json
import time
import pandas as pd
from datetime import datetime
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

# === Load session ===
SESSION_FILE = "selenium_session.json"

def load_session():
    try:
        with open(SESSION_FILE, "r") as f:
            session_info = json.load(f)
        return session_info
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error loading session: {e}")
        return None

session_info = load_session()
if not session_info:
    print("Session not found. Please run login.py first.")
    exit()

chrome_service = Service("C:\\Drivers\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe")
chrome_options = Options()
chrome_options.add_argument("--start-minimized")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

try:
    driver = webdriver.Remote(
        command_executor=session_info['executor_url'],
        options=chrome_options
    )
    driver.session_id = session_info['session_id']
    print("✅ Reusing existing browser session.")
except WebDriverException as e:
    print(f"❌ Failed to reuse session: {e}. Please log in again by running login.py.")
    exit()

# === Load EMPLIDs ===
csv_path = "D:\\Xavier University\\CISO\\Scripts\\Selenium\\SLMIS\\Account Creation\\accounts.csv"
emplid_df = pd.read_csv(csv_path)
student_ids = emplid_df["EMPLID"].dropna().astype(str).tolist()

success_log = []
error_log = []

pbar = tqdm(student_ids, desc="Processing Accounts", ncols=100)

for emplid in pbar:
    generated_password = None

    # === 1) Open User Profile page ===
    driver.get("https://slmis.xu.edu.ph/psp/ps/EMPLOYEE/HRMS/c/MAINTAIN_SECURITY.USERMAINT.GBL?PORTALPARAM_PTCNAV=PT_USERMAINT_GBL&EOPP.SCNode=HRMS&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=PT_USER_PROFILES&EOPP.SCLabel=User%20Profiles&EOPP.SCPTfname=PT_USER_PROFILES&FolderPath=PORTAL_ROOT_OBJECT.PT_PEOPLETOOLS.PT_SECURITY.PT_USER_PROFILES.PT_USERMAINT_GBL&IsFolder=false")
    time.sleep(2)

    try:
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "TargetContent")))

        # Enter EMPLID
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "PSOPRDEFN_SRCH_OPRID"))
        )
        search_box.clear()
        search_box.send_keys(emplid)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        # Check if 'No matching values were found.'
        no_result = driver.find_elements(By.XPATH, "//h2[contains(text(), 'No matching values were found.')]")

        if no_result:

            # === 1) Open Save As page ===
            driver.get("https://slmis.xu.edu.ph/psp/ps/EMPLOYEE/HRMS/c/MAINTAIN_SECURITY.USER_SAVEAS.GBL?FolderPath=PORTAL_ROOT_OBJECT.PT_PEOPLETOOLS.PT_SECURITY.PT_USER_PROFILES.PT_USER_SAVEAS_GBL&IsFolder=false&IgnoreParamTempl=FolderPath%2cIsFolder")
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "TargetContent")))

            # === 2) Input static ID ===
            static_id_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "PSOPRDEFN_SRCH_OPRID"))
            )
            static_id_input.clear()
            static_id_input.send_keys("61218")
            static_id_input.send_keys(Keys.ENTER)
            time.sleep(2)

            # === 3) Input new EMPLID ===
            clone_oprid_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "DERIVED_CLONE_OPRID"))
            )
            clone_oprid_input.clear()
            clone_oprid_input.send_keys(emplid)

            # === 4) Input static password in both fields ===
            pw1 = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "DERIVED_CLONE_OPERPSWD"))
            )
            pw2 = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "DERIVED_CLONE_OPERPSWDCONF"))
            )
            pw1.clear()
            pw2.clear()
            pw1.send_keys("dgPX4r2w")
            pw2.send_keys("dgPX4r2w")

            # === 5) Click Save ===
            save_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "#ICSave"))
            )
            save_btn.click()
            print(f"✅ Saved new user {emplid} by cloning 24440")
            time.sleep(3)

            # === 6) Go to ID tab ===
            id_tab = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ICTAB_1"))
            )
            id_tab.click()
            time.sleep(2)

            # === 7) Ensure alias type EMP ===
            alias_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "PSOPRALIAS_OPRALIASTYPE$0"))
            )
            alias_dropdown.send_keys("EMP")

            for attempt in range(5):
                try:
                    alias_value_input = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "PSOPRALIAS_WRK_ATTRVALUE$0"))
                    )
                    # Focus it and type
                    alias_value_input.click()
                    alias_value_input.clear()
                    alias_value_input.send_keys(emplid)
                    alias_value_input.send_keys(Keys.RETURN)
                    print(f"✅ Entered alias value {emplid}")
                    break

                except StaleElementReferenceException:
                    print("♻️ Input went stale — retrying...")
                    time.sleep(1)


            # === 8) Click Set Description ===
            for attempt in range(5):
                try:
                    set_descr = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "PSUSRPRFL_WRK_SET_DESCR_PB$0"))
                    )
                    set_descr.click()
                    print("✅ Clicked Set Description")
                    break

                except StaleElementReferenceException:
                    print("♻️ Set Description link was stale — retrying...")
                    time.sleep(1)

            # === 9) Save ===
            save_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "#ICSave"))
            )
            save_btn.click()
            print("✅ Saved description")
            time.sleep(2)

            # === 10) Return to General tab ===
            general_tab = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ICTAB_0"))
            )
            general_tab.click()
            time.sleep(1)

            name_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "PSOPRDEFN_OPRDEFNDESC"))
            )

            full_name = name_element.text.strip()


            # === 11) Click Generate ===
            generate_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "PSOPRDEFN_OPRTYPE"))
            )
            generate_btn.click()
            time.sleep(2)

            # === 12) Get new password ===
            generated_password = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "PSUSRPRFL_WRK_DESCR"))
            ).text.strip()

            # === Final Save ===
            save_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "#ICSave"))
            )
            save_btn.click()

            

            success_log.append({
                "Name": full_name,
                "EMPLID": emplid,
                "Password": generated_password,
                "Status": "New Account",
            })

            continue


        print(f"✅ {emplid} found — updating profile")
        # === Get full name ===
        name_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "PSOPRDEFN_OPRDEFNDESC"))
        )
        full_name = name_element.text.strip()
        print(f"👤 Name: {full_name}")

        # === Check & untick Password Expired ===
        pwd_expired_chk = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "PSUSRPRFL_WRK_PSWDEXPIRED"))
        )
        if pwd_expired_chk.get_attribute("checked"):
            pwd_expired_chk.click()
            print(f"✅ Unticked Password Expired for {emplid}")
            time.sleep(2)
        else:
            print(f"✔️ Checkbox not checked for {emplid}")

        # === Click Generate ===
        generate_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "PSOPRDEFN_OPRTYPE"))
        )
        generate_btn.click()
        print(f"🔑 Clicked Generate for {emplid}")
        time.sleep(2)

        # === Get generated password ===
        generated_password = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "PSUSRPRFL_WRK_DESCR"))
        ).text.strip()
        print(f"✅ New password: {generated_password}")

        # === Click Save ===
        save_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "#ICSave"))
        )
        save_btn.click()
        print(f"💾 Saved profile for {emplid}")
        time.sleep(2)

        # === 2) Open Gmail Account creation ===
        driver.get("https://slmis.xu.edu.ph/psp/ps/EMPLOYEE/HRMS/c/XU_GMAIL_ACCOUNT.XU_GMAIL_ACCOUNT.GBL")
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "TargetContent")))

        # Input EMPLID
        gmail_id_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "XU_GMAIL_EMPLID"))
        )
        gmail_id_input.clear()
        gmail_id_input.send_keys(emplid)
        gmail_id_input.send_keys(Keys.ENTER)
        time.sleep(2)

        # Input password
        gmail_pw_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "XU_GMAIL_PASSWORD2"))
        )
        gmail_pw_input.clear()
        gmail_pw_input.send_keys(generated_password)
        time.sleep(1)

        # Click Save
        gmail_save = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "#ICSave"))
        )
        gmail_save.click()
        print(f"📧 Gmail created for {emplid}")

        success_log.append({
            "Name": full_name,
            "EMPLID": emplid,
            "Password": generated_password,
            "Status": "Exists"
        })

    except Exception as e:
        print(f"❌ Error with {emplid}: {e}")
        error_log.append({
            "EMPLID": emplid,
            "Error": str(e)
        })

# === Report ===
success_df = pd.DataFrame(success_log)
error_df = pd.DataFrame(error_log)
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_file = f"account_creation_report_{timestamp}.xlsx"

with pd.ExcelWriter(output_file) as writer:
    success_df.to_excel(writer, sheet_name="Success", index=False)
    error_df.to_excel(writer, sheet_name="Errors", index=False)

print(f"\n✅ All done! Report saved as: {output_file}")
