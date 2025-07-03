# Selenium Login Automation for SLMIS

This project automates the login process for the [SLMIS platform](https://slmis.xu.edu.ph/psp/ps/?cmd=login) using Selenium WebDriver. Once you provide your credentials, it saves the session data for later use.

## Features

- Automated login process for SLMIS platform.
- Saves session data (session_id, executor_url) for future use.
- Runs with Chrome WebDriver.
- Can be easily converted into an executable.

## Prerequisites

Before running this script, make sure you have the following installed:

- Python 3.x
- Chrome WebDriver (Make sure you download the correct version based on your Chrome version)
- Chrome driver is attached in the file already, change directory referencing based on your directory

### Required Python Libraries



You will need to install the following Python packages:

- `selenium`
- `getpass`
- `json`

To install the necessary libraries, run the following command in your terminal:

### Optional
```bash
pip install pyinstaller

```bash
pip install -r requirements.txt

run python login.py or py login.py

