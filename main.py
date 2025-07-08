import pandas as pd
import pywhatkit
import time
import pyautogui
from pathlib import Path
from pywhatkit.core import exceptions

def format_phone_number(phone: str, country_code: str) -> str:
    """
    Formats a phone number by ensuring it has a country code prefix.
    
    Args:
        phone: The phone number as a string.
        country_code: The country code to add (e.g., '+91').
        
    Returns:
        The formatted phone number string.
    """
    phone = str(phone).strip()
    if not phone.startswith('+'):
        phone = f"{country_code}{phone}"
    return phone

def send_whatsapp_messages(excel_file: str, column_name: str, message: str, country_code: str) -> tuple[int, int]:
    """
    Reads phone numbers from an Excel file and sends a WhatsApp message to each.
    
    Args:
        excel_file: Path to the Excel file.
        column_name: The name of the column containing phone numbers.
        message: The message to send.
        country_code: The default country code for formatting.
        
    Returns:
        A tuple containing the count of successful and failed messages.
    """
    success_count = 0
    failure_count = 0
    
    try:
        df = pd.read_excel(excel_file)
        if column_name not in df.columns:
            print(f"❌ Error: Column '{column_name}' not found in the Excel file.")
            return 0, 0
            
        phone_numbers = df[column_name].dropna().tolist()
        total_messages = len(phone_numbers)
        print(f"\nFound {total_messages} phone numbers. Starting the process...")

        for i, phone in enumerate(phone_numbers, 1):
            formatted_phone = format_phone_number(phone, country_code)
            print(f"🔄 [{i}/{total_messages}] Preparing to send to: {formatted_phone}")
            
            try:
                pywhatkit.sendwhatmsg_instantly(
                    phone_no=formatted_phone,
                    message=message,
                    tab_close=True,
                    close_time=3 # Time in seconds before the tab is closed
                )
                
                # pyautogui is a fallback in case tab_close fails
                time.sleep(4) # Give time for the message to be sent
                pyautogui.hotkey('ctrl', 'w')
                
                print(f"✅ Message sent successfully to {formatted_phone}")
                success_count += 1

            except exceptions.CountryCodeException:
                print(f"❌ Failed: Invalid phone number or country code for {formatted_phone}")
                failure_count += 1
            except Exception as e:
                print(f"❌ Failed to send to {formatted_phone}. Error: {e}")
                failure_count += 1

            # A longer delay to reduce the risk of being flagged as spam
            print("--- Waiting for next message ---")
            time.sleep(10)

    except FileNotFoundError:
        print(f"❌ Error: The file '{excel_file}' was not found.")
    except Exception as e:
        print(f"❌ An unexpected error occurred: {e}")
        
    return success_count, failure_count

def main():
    """Main function to drive the script."""
    print("--- WhatsApp Bulk Messenger ---")

    # 1. Get Excel file path
    while True:
        excel_file = input("Enter the full path to your Excel file: ").strip('"')
        if Path(excel_file).is_file() and excel_file.lower().endswith(('.xlsx', '.xls')):
            break
        else:
            print("Invalid file path. Please provide a valid path to an .xlsx or .xls file.")

    # 2. Get column name
    column_name = input("Enter the column name with phone numbers (default: PhoneNumbers): ") or 'PhoneNumbers'
    
    # 3. Get default country code
    while True:
        country_code = input("Enter the default country code (e.g., +91 for India): ").strip()
        if country_code.startswith('+') and country_code[1:].isdigit():
            break
        else:
            print("Invalid format. Please enter a code like '+1', '+44', '+91', etc.")
            
    # 4. Get message
    print("Enter the message you want to send. You can use multiple lines. Press Enter on an empty line to finish.")
    message_lines = []
    while True:
        line = input()
        if line == "":
            break
        message_lines.append(line)
    message = "\n".join(message_lines)

    if not message:
        print("Message cannot be empty. Exiting.")
        return

    # 5. Confirm and Execute
    print("\n--- Please Review ---")
    print(f"File: {excel_file}")
    print(f"Column: {column_name}")
    print(f"Country Code: {country_code}")
    print(f"Message:\n{message}")
    
    confirm = input("\nAre you sure you want to proceed? (yes/no): ").lower()
    if confirm == 'yes':
        success, failed = send_whatsapp_messages(excel_file, column_name, message, country_code)
        
        print("\n--- Execution Summary ---")
        print(f"✅ Successful messages: {success}")
        print(f"❌ Failed messages: {failed}")
        print("--------------------------")
    else:
        print("Operation cancelled by user.")

if __name__ == "__main__":
    main()
    print("Script execution completed.")