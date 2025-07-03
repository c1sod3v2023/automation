# Selenium Login Automation for SLMIS

This project automates the login process for the [SLMIS platform](https://slmis.xu.edu.ph/psp/ps/?cmd=login) using Selenium WebDriver. Once you provide your credentials, it saves the session data for later use.

## Features

- Automated login process for SLMIS platform.
- Saves session data (session_id, executor_url) for future use.
- Runs with Chrome WebDriver.
- Can be easily converted into an executable.

## Prerequisites

Before running this script, make sure you have the following installed:

- **Python 3.x**: Download and install Python from [here](https://www.python.org/downloads/).
- **Chrome WebDriver**: Ensure you have the correct version of Chrome WebDriver matching your installed Chrome version. You can download it from [here](https://sites.google.com/a/chromium.org/chromedriver/).
  
  *Once downloaded, extract the WebDriver and ensure the path is referenced correctly in the script. The script includes a default path, but you can change it according to your system.*

### Setting up Dependencies

To run this script, you need to install some required Python libraries. You can do this easily by following the steps below:

### Installing Python Libraries

1. **Install Python**:
   - Download and install Python from the official website: [https://www.python.org/downloads/](https://www.python.org/downloads/).
   - Make sure to check the box that says **Add Python to PATH** during installation.

2. **Install the required Python packages**:

   In your terminal (Command Prompt on Windows, Terminal on macOS/Linux), run the following command to install all required dependencies:

   ```bash
   pip install selenium pyinstaller

   pip install -r requirements.txt

   python login.py

   python matriculate.py

   python account_creation.py

---

### Key Sections in the README:

1. **Installation of Python & Chrome WebDriver**: The installation steps for Python and Chrome WebDriver are clearly described, along with where to download them.
   
2. **Installing Dependencies**: Instructions for installing Python dependencies (`selenium`, `pyinstaller`) with both a single command and using a `requirements.txt` file.

3. **Running the Script**: A clear explanation of how to run the Python script after setting everything up.

4. **Optional Executable Generation**: Instructions for converting the script into a standalone `.exe` file using PyInstaller.

5. **Troubleshooting**: A brief troubleshooting section to help users resolve common issues like mismatched Chrome WebDriver versions and problems with PyInstaller.

---

With this complete README, users can easily set up the environment, run the script, or even convert it into an executable without needing to install Python or dependencies themselves.
