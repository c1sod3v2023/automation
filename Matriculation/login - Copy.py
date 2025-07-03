import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

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

def get_chrome_driver_path():
    """Get the correct path to chromedriver depending on the environment."""
    # If the app is running as a bundled executable
    if getattr(sys, 'frozen', False):
        # _MEIPASS is a temporary folder where PyInstaller stores the bundled files
        chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
    else:
        # If it's running from source, use the relative path
        chromedriver_path = os.path.join(os.path.dirname(__file__), 'chromedriver.exe')

    return chromedriver_path

def main():
    """Main function to perform the login and save session."""
    
    # Get the correct chromedriver path (this will work for both bundled & non-bundled)
    chrome_driver_path = get_chrome_driver_path()

    # Start the Chrome driver
    driver = setup_chrome_driver(chrome_driver_path)

    # Keep prompting until successful login
    while True:
        username = input("Enter your SLMIS username: ")
        password = input("Enter your password (hidden): ")

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
