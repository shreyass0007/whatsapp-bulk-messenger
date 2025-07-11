# WhatsApp Bulk Messenger

A Python script to send WhatsApp messages in bulk using data from an Excel file. This tool leverages `pandas`, `pywhatkit`, and `pyautogui` to automate message sending via WhatsApp Web.

## Features
- Send personalized or bulk messages to multiple contacts via WhatsApp Web
- Read phone numbers from an Excel file (`.xlsx` or `.xls`)
- Automatic phone number formatting with country code
- Interactive command-line interface for user input
- Error handling and summary reporting

## Requirements
- Python 3.7+
- Google Chrome (or your default browser)
- WhatsApp Web account (logged in)

### Python Packages
- pandas
- pywhatkit
- pyautogui
- openpyxl (for `.xlsx` support)
- xlrd (for `.xls` support)

Install all dependencies with:
```bash
pip install pandas pywhatkit pyautogui openpyxl xlrd
```

## Setup
1. Clone this repository or download the script.
2. Prepare your Excel file with a column containing phone numbers (default column name: `PhoneNumbers`).
3. Ensure you are logged into WhatsApp Web in your default browser.

## Usage
Run the script from your terminal:
```bash
python main.py
```
Follow the prompts:
- Enter the path to your Excel file
- Enter the column name with phone numbers (or press Enter for default)
- Enter the country code (e.g., `+91`)
- Enter your message (multi-line supported)
- Confirm to start sending messages

**Note:**
- Do not use your computer while the script is running, as it automates browser actions.
- There is a delay between messages to avoid spam detection.
- Make sure your Excel file is closed before running the script.

## Excel File Format

Your Excel file should contain a column with phone numbers. By default, the script looks for a column named `PhoneNumbers`, but you can specify a different column name when prompted.

**Example Excel Table:**

| PhoneNumbers   |
|---------------|
| 9876543210    |
| 9123456789    |
| 9988776655    |

- The column can be named anything, but you must enter the exact name when running the script.
- Phone numbers can be with or without the country code. If missing, the script will prepend the country code you provide.
- The file must be in `.xlsx` or `.xls` format.

## Disclaimer
This tool is for educational and personal use only. Use responsibly and respect WhatsApp's terms of service.

## License
MIT License
