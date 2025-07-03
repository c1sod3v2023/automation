import time
import json
import getpass
import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# Path to store session data
SESSION_FILE = "selenium_session.json"

def setup_chrome_driver(chrome_driver_path: str) -> webdriver.Chrome:
    """Sets up the Chrome driver with the necessary options."""
    chrome_service = Service(chrome_driver_path)

    # Configure Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    return webdriver.Chrome(service=chrome_service, options=chrome_options)

def save_session(driver: webdriver.Chrome) -> None:
    """Saves the session info (session_id, executor_url) to a file."""
    session_info = {
        "session_id": driver.session_id,
        "executor_url": driver.command_executor._url
    }
    with open(SESSION_FILE, "w") as f:
        json.dump(session_info, f)

def login_to_site(driver: webdriver.Chrome, username: str, password: str) -> bool:
    """Logs in to the SLMIS platform and returns True if successful, False otherwise."""
    driver.get("https://slmis.xu.edu.ph/psp/ps/?cmd=login")

    try:
        username_element = driver.find_element(By.ID, "userid")
        password_element = driver.find_element(By.ID, "pwd")
        login_btn_element = driver.find_element(By.CLASS_NAME, "ps-button")

        username_element.send_keys(username)
        password_element.send_keys(password)
        login_btn_element.click()

        time.sleep(3)  # Allow page to load

        # Check for login failure by examining the URL or title
        if "Sign In" in driver.title or "login" in driver.current_url:
            print("❌ Login failed. Please check your credentials.")
            return False

        print("✅ Login successful.")
        return True

    except Exception as e:
        print(f"⚠️ Error during login: {e}")
        return False

def main():
    """Main function to perform the login and save session."""
    # Determine the correct path to the Chrome WebDriver
    if getattr(sys, 'frozen', False):
        # Running as an executable
        chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver')
    else:
        # Running as a Python script
        chromedriver_path = "C:\\path\\to\\chromedriver.exe"  # Replace with your actual WebDriver path

    # Start the Chrome driver
    driver = setup_chrome_driver(chromedriver_path)

    # Keep prompting until successful login
    while True:
        username = input("Enter your SLMIS username: ")
        password = getpass.getpass("Enter your password (hidden): ")

        if login_to_site(driver, username, password):
            save_session(driver)
            print("💾 Session saved. You can now run matriculate.py to continue.")
            break
        else:
            print("🔁 Try logging in again.\n")

    # Keep browser open until manually closed
    try:
        while True:
            time.sleep(100)
    except KeyboardInterrupt:
        print("🔚 Script interrupted. Closing browser.")
        driver.quit()

if __name__ == "__main__":
    main()
