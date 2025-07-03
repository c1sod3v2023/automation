import os
import sys
import time
import getpass
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# Enable logging for debugging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def setup_chrome_driver(chrome_driver_path: str) -> webdriver.Chrome:
    """Sets up the Chrome driver with the necessary options."""
    logging.debug(f"Setting up ChromeDriver with path: {chrome_driver_path}")
    try:
        chrome_service = Service(chrome_driver_path)
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
        logging.info("ChromeDriver successfully initialized.")
        return driver
    except Exception as e:
        logging.error(f"Error initializing ChromeDriver: {e}")
        raise

def get_chrome_driver_path():
    """Get the correct path to chromedriver depending on the environment."""
    if getattr(sys, 'frozen', False):
        # _MEIPASS is a temporary folder where PyInstaller stores the bundled files
        chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
    else:
        chromedriver_path = os.path.join(os.path.dirname(__file__), 'chromedriver.exe')

    logging.debug(f"Chromedriver path is: {chromedriver_path}")
    return chromedriver_path

def login_to_site(driver: webdriver.Chrome, username: str, password: str) -> bool:
    """Logs into the SLMIS platform and returns True if successful, False otherwise."""
    logging.debug("Attempting login...")
    try:
        driver.get("https://slmis.xu.edu.ph/psp/ps/?cmd=login")

        username_element = driver.find_element(By.ID, "userid")
        password_element = driver.find_element(By.ID, "pwd")
        login_btn_element = driver.find_element(By.CLASS_NAME, "ps-button")

        username_element.send_keys(username)
        password_element.send_keys(password)
        login_btn_element.click()

        time.sleep(3)  # Allow page to load

        if "Sign In" in driver.title or "login" in driver.current_url:
            logging.warning("Login failed. Incorrect credentials or login page.")
            return False

        logging.info("Login successful!")
        return True
    except Exception as e:
        logging.error(f"Error during login: {e}")
        return False

def save_session(driver: webdriver.Chrome) -> None:
    """Saves the session info (session_id, executor_url) to a file."""
    try:
        session_info = {
            "session_id": driver.session_id,
            "executor_url": driver.command_executor._url
        }
        with open("selenium_session.json", "w") as f:
            json.dump(session_info, f)
        logging.info("Session saved successfully.")
    except Exception as e:
        logging.error(f"Error saving session: {e}")

def main():
    """Main function to perform the login and save session."""
    try:
        # Get the correct chromedriver path
        chrome_driver_path = get_chrome_driver_path()

        # Start the Chrome driver
        driver = setup_chrome_driver(chrome_driver_path)

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
    except Exception as e:
        logging.error(f"Critical error: {e}")

if __name__ == "__main__":
    main()
