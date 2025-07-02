import json
import time
import pandas as pd
from datetime import datetime
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === GUI Class ===
class ProgressBarGUI:
    def __init__(self, total, start_processing_callback):
        self.total = total
        self.current = 0
        self.callback = start_processing_callback

        self.root = tk.Tk()
        self.root.title("Matriculation Progress")
        self.root.geometry("400x260")

        # Input Fields
        tk.Label(self.root, text="Institution Code (e.g., XUNV):").pack(pady=(10, 0))
        self.institution_entry = tk.Entry(self.root)
        self.institution_entry.pack()

        tk.Label(self.root, text="Career Code (e.g., UGRD):").pack(pady=(10, 0))
        self.career_entry = tk.Entry(self.root)
        self.career_entry.pack()

        self.start_button = tk.Button(self.root, text="Start Processing", command=self.start_callback)
        self.start_button.pack(pady=10)

        self.label = tk.Label(self.root, text="")
        self.label.pack(pady=5)

        self.progress = ttk.Progressbar(self.root, length=350, mode='determinate', maximum=total)
        self.progress.pack(pady=5)

        self.status_label = tk.Label(self.root, text="")
        self.status_label.pack(pady=5)

    def start_callback(self):
        institution = self.institution_entry.get().strip()
        career = self.career_entry.get().strip()

        if not institution or not career:
            messagebox.showerror("Missing Input", "Please enter both Institution and Career codes.")
            return

        self.institution_entry.config(state="disabled")
        self.career_entry.config(state="disabled")
        self.start_button.config(state="disabled")
        self.label.config(text="Processing students...")

        threading.Thread(target=self.callback, args=(self.update, self.done, institution, career)).start()

    def update(self, student_id):
        self.current += 1
        self.progress['value'] = self.current
        self.status_label.config(text=f"Processed EMPLID: {student_id}")
        self.root.update_idletasks()

    def done(self):
        self.status_label.config(text="✅ Processing complete!")
        self.root.after(3000, self.root.destroy)

    def run(self):
        self.root.mainloop()

# === Main Processing Function ===
def start_processing(update_gui, finish_gui, institution_input, career_input):
    SESSION_FILE = "selenium_session.json"
    success_log = []
    error_log = []

    def load_session():
        try:
            with open(SESSION_FILE, "r") as f:
                return json.load(f)
        except:
            print("❌ Session not found. Please run login.py first.")
            return None

    session_info = load_session()
    if not session_info:
        return

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
        print(f"❌ Failed to reuse session: {e}")
        return

    # CSV File
    csv_path = "D:\\Xavier University\\CISO\\Scripts\\Selenium\\SLMIS\\Matriculation\\matriculate.csv"
    emplid_df = pd.read_csv(csv_path)
    student_ids = emplid_df["EMPLID"].dropna().astype(int).tolist()

    for student_id in student_ids:
        try:
            update_gui(student_id)

            driver.get("https://slmis.xu.edu.ph/psp/ps/EMPLOYEE/HRMS/c/PROCESS_APPLICATIONS.ADM_APPL_MAINTNCE.GBL")
            time.sleep(1)
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

            actions = ActionChains(driver)
            actions.key_down(Keys.ALT).send_keys("1").key_up(Keys.ALT).perform()
            time.sleep(1)

            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ICTAB_3"))).click()

            prog_action_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ADM_APPL_PROG_PROG_ACTION$0")))
            prog_action_value = prog_action_input.get_attribute("value").strip()

            if prog_action_value == "ADMT":
                try:
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "#ICCorrection"))).click()
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ADM_APPL_PROG$new$0$$0"))).click()
                    matr_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "ADM_APPL_PROG_PROG_ACTION$0")))
                    matr_field.clear()
                    matr_field.send_keys("MATR")
                    matr_field.send_keys(Keys.ENTER)

                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "DERIVED_ADM_CREATE_PROG_PB$0"))).click()
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.NAME, "#ICSave"))).click()

                    success_log.append({"Student ID": student_id, "Program Action": prog_action_value})
                except Exception as inner_err:
                    error_log.append({"Student ID": student_id, "Error": str(inner_err), "Program Action": prog_action_value})
            else:
                error_log.append({"Student ID": student_id, "Error": "Program Action is not ADMT", "Program Action": prog_action_value})

        except Exception as e:
            error_log.append({"Student ID": student_id, "Error": str(e), "Program Action": "N/A"})

    # Save logs
    success_df = pd.DataFrame(success_log)
    error_df = pd.DataFrame(error_log)
    timestamp = datetime.now().strftime("%d_%m_%H_%M")
    filename = f"matriculation_report_{timestamp}_{career_input.upper()}.xlsx"
    with pd.ExcelWriter(filename) as writer:
        success_df.to_excel(writer, sheet_name="Success", index=False)
        error_df.to_excel(writer, sheet_name="Errors", index=False)

    print(f"\n✅ Excel report saved as {filename}")
    finish_gui()

# === Run GUI ===
if __name__ == "__main__":
    csv_path = "D:\\Xavier University\\CISO\\Scripts\\Selenium\\SLMIS\\Matriculation\\matriculate.csv"
    emplid_df = pd.read_csv(csv_path)
    student_ids = emplid_df["EMPLID"].dropna().astype(int).tolist()
    gui = ProgressBarGUI(len(student_ids), start_processing)
    gui.run()
