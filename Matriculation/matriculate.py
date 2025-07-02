import json
import time
import pandas as pd
from datetime import datetime
from tqdm import tqdm  # Terminal progress bar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

# === Load EMPLIDs from CSV ===
csv_path = "D:\\Xavier University\\CISO\\Scripts\\Selenium\\SLMIS\\Matriculation\\matriculate.csv"
emplid_df = pd.read_csv(csv_path)
student_ids = emplid_df["EMPLID"].dropna().astype(int).tolist()

# === Logs ===
success_log = []
error_log = []

# === User Input ===
institution_input = input("Enter Institution code (e.g., XUNV): ").strip()
career_input = input("Enter Career code (e.g., UGRD): ").strip()

# === Progress bar ===
pbar = tqdm(student_ids, desc="Matriculating students", ncols=100)

for student_id in pbar:
    pbar.set_postfix_str(f"EMPLID: {student_id}")
    try:
        driver.get("https://slmis.xu.edu.ph/psp/ps/EMPLOYEE/HRMS/c/PROCESS_APPLICATIONS.ADM_APPL_MAINTNCE.GBL")
        time.sleep(2)
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "TargetContent")))

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ADM_MAINT_SCTY_EMPLID$op")))
        for option in driver.find_element(By.ID, "ADM_MAINT_SCTY_EMPLID$op").find_elements(By.TAG_NAME, "option"):
            if option.get_attribute("value") == "2":
                option.click()
                break

        driver.find_element(By.NAME, "ADM_MAINT_SCTY_EMPLID").send_keys(str(student_id))
        time.sleep(1)

        institution_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "ADM_MAINT_SCTY_INSTITUTION")))
        institution_field.clear()
        institution_field.send_keys(institution_input)
        time.sleep(1)

        career_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "ADM_MAINT_SCTY_ACAD_CAREER")))
        career_field.clear()
        career_field.send_keys(career_input)
        career_field.send_keys(Keys.TAB)

        time.sleep(1)
        actions = ActionChains(driver)
        actions.key_down(Keys.ALT).send_keys("1").key_up(Keys.ALT).perform()
        time.sleep(1)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ICTAB_3"))).click()

        prog_action_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ADM_APPL_PROG_PROG_ACTION$0")))
        prog_action_value = prog_action_input.get_attribute("value").strip()

        if prog_action_value == "ADMT":
            try:
                pbar.set_postfix_str(f"{student_id} → Saving...")
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "#ICCorrection"))).click()
                time.sleep(3)

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ADM_APPL_PROG$new$0$$0"))).click()
                time.sleep(3)

                matr_field = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ADM_APPL_PROG_PROG_ACTION$0")))
                matr_field.clear()
                matr_field.send_keys("MATR")
                matr_field.send_keys(Keys.ENTER)
                time.sleep(2)

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "DERIVED_ADM_CREATE_PROG_PB$0"))).click()
                time.sleep(2)

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.NAME, "#ICSave"))).click()

                success_log.append({
                    "Student ID": student_id,
                    "Program Action": prog_action_value
                })

            except Exception as inner_err:
                pbar.set_postfix_str(f"{student_id} ❌ Failed to save")
                error_log.append({
                    "Student ID": student_id,
                    "Error": str(inner_err),
                    "Program Action": prog_action_value
                })
        else:
            pbar.set_postfix_str(f"{student_id} ⚠️ Already MATR")
            error_log.append({
                "Student ID": student_id,
                "Error": "Program Action is already MATR",
                "Program Action": prog_action_value
            })

    except Exception as outer_err:
        pbar.set_postfix_str(f"{student_id} ❌ Error occurred")
        error_log.append({
            "Student ID": student_id,
            "Error": str(outer_err),
            "Program Action": "N/A"
        })

# === Finish ===
pbar.set_postfix_str("✔ Done")
pbar.close()
print("\n✅ Automation complete.")

# === Export to Excel ===
success_df = pd.DataFrame(success_log)
error_df = pd.DataFrame(error_log)

# All Processed Sheet
processed_log = []
for student_id in student_ids:
    status = "✅ Success" if any(s["Student ID"] == student_id for s in success_log) else \
             "❌ Error" if any(e["Student ID"] == student_id for e in error_log) else \
             "⚠️ Unknown"
    processed_log.append({
        "Student ID": student_id,
        "Status": status
    })
processed_df = pd.DataFrame(processed_log)

timestamp = datetime.now().strftime("%d_%m_%H_%M")
output_filename = f"matriculation_report_{timestamp}_{career_input.upper()}.xlsx"

with pd.ExcelWriter(output_filename) as writer:
    success_df.to_excel(writer, sheet_name="Success", index=False)
    error_df.to_excel(writer, sheet_name="Errors", index=False)
    processed_df.to_excel(writer, sheet_name="All Processed", index=False)

print(f"\n📁 Excel report saved as: {output_filename}")
