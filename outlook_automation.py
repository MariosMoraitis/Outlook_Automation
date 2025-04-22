from datetime import datetime
import win32com.client
import json
import os
try:
    # Load settings from settings.json
    with open('settings.json', 'r', encoding='utf-8') as f:
        settings = json.load(f)
        
     # Check if all required keys exist in the settings
    required_keys: list[str] = ["lang", "signature"]
    for key in required_keys:
        if key not in settings:
            raise KeyError(f"Missing key '{key}' in JSON file.")
    
    # Dynamically import language-specific config based on selected language
    if settings["lang"].strip().lower() == "gr":
        from config_gr import *
    else:
        from config_en import *

except KeyError as e:
     # Handle missing required keys in settings.json
    print(f"\n⚠️ ERROR: {e}")
    quit()
except FileNotFoundError:
    # Handle missing settings.json file
    print("\n⚠️ ERROR: JSON file: '....json' not found!!!.")
    quit()
except Exception as e:
    # Handle any other unexpected errors
    print("\n⚠️ ERROR: Unexpected error occurred...")
    quit()



def prepare_body(body, user_name) -> str:
    """
    Constructs the full email body with a greeting and a sign-off,
    based on the current time of day and the provided user name.

    Args:
        body (str): The main message body of the email.
        user_name (str): The name of the sender to be included in the signature.

    Returns:
        str: The final formatted email body.
    """
    time: int = datetime.now().hour
    if time < 12:
        greeting: str = f"{MORNING_GREETING},\n\n"
    else:
        greeting: str = f"{EVENING_GREETING},\n\n"
    
    sign_off: str = f'\n\n{SIGN_OFF},\n{user_name}'

    return greeting + body + sign_off

def load_json() -> dict | None:
    """
    Loads the email configuration from 'mail_config.json'.

    Returns:
        dict | None: A dictionary with the email data if successful,
        or None if the file is missing or an error occurs.
    """
    try:
        with open('mail_config.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data
    
    except FileNotFoundError:
        print("\n⚠️ ERROR: JSON file: '....json' not found!!!.")
    except Exception as e:
        print("\n⚠️ ERROR: Unexpected error occurred...")

def outlook_main() -> None:
    """
    Main function that prepares and opens an Outlook email draft.
    It loads user input and email settings, formats the message,
    and uses Outlook's COM interface to open the email ready to send.
    """
    issue_number: str = input('Enter the issue number: ').strip().upper()
    print("Let's prepare the mail...")

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    data = load_json()
    if data:
        mail.to = data["recipients"]
        mail.cc = data["cc_recipients"]
        mail.Subject = f"{issue_number}"

        data["body"] += f'{issue_number}.'
        body: str = prepare_body(data["body"], data["user_name"])
        mail.Body = body

        mail.Display()

if __name__ == '__main__':
    print('Welcome to Outlook automation!')
    outlook_main()
    os.system('pause')