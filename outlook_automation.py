from datetime import datetime
import win32com.client
import json
import importlib

def load_settings_and_config():
    try:
        # Load settings from settings.json
        with open('parameters/settings.json', 'r', encoding='utf-8') as f:
            settings = json.load(f)
            
        # Check if all required keys exist in the settings
        required_keys: list[str] = ["lang", "signature"]
        for key in required_keys:
            if key not in settings:
                raise KeyError(f"Missing key '{key}' in JSON file.")
        
        # Dynamically import language-specific config based on selected language
        if settings["lang"].strip().lower() == "gr":
            import config_gr
            importlib.reload(config_gr)
            config = config_gr
        else:
            import config_en
            importlib.reload(config_en)
            config = config_en

        signature_flag = False
        if settings["signature"].strip().lower() == 'yes':
            signature_flag = True
        
        return config, signature_flag

    except KeyError as e:
        # Handle missing required keys in settings.json
        print(f"\n⚠️ ERROR: {e}")
        quit()
    except FileNotFoundError:
        # Handle missing settings.json file
        print("\n⚠️ ERROR: JSON file: 'settings.json' not found!!!.")
        quit()
    except Exception as e:
        # Handle any other unexpected errors
        print("\n⚠️ ERROR: Unexpected error occurred...")
        quit()

def get_greeting(config) -> str:
    """
    Constructs a greeting, based on the current time of day.

    Returns:
        str: The final formatted greeting.
    """
    time: int = datetime.now().hour
    if time < 12:
        greeting: str = f"{config.MORNING_GREETING}," # type: ignore
    else:
        greeting: str = f"{config.EVENING_GREETING}," # type: ignore
    
    return greeting

def load_json() -> dict | None:
    """
    Loads the email configuration from 'mail_config.json'.

    Returns:
        dict | None: A dictionary with the email data if successful,
        or None if the file is missing or an error occurs.
    """
    try:
        with open('parameters/mail_config.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data
    
    except FileNotFoundError:
        print("\n⚠️ ERROR: JSON file: mail_config.json' not found!!!.")
    except Exception as e:
        print("\n⚠️ ERROR: Unexpected error occurred...")

def outlook_main(issue_number) -> None:
    """
    Main function that prepares and opens an Outlook email draft.
    It loads user input and email settings, formats the message,
    and uses Outlook's COM interface to open the email ready to send.
    """
    config, signature_flag = load_settings_and_config()
    if not config or signature_flag is None:
        return
    
    print("Let's prepare the mail...")

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    data = load_json()
    if data:
        mail.to = data["recipients"]
        mail.cc = data["cc_recipients"]
        if data["subject"]:
            mail.Subject = data["subject"]
        else:
            mail.Subject = f"{issue_number}"

        _greeting: str = get_greeting(config)
        body: str = f'{data["body"]} {issue_number}.'
        mail.Display()
        if signature_flag:
            signature = mail.HTMLBody  # This now contains the default Outlook signature
            # Now insert your text above the signature
            final_body: str = f'{_greeting}<br><br>{body}\n\n{signature}'
            mail.HTMLBody = final_body
        else:
            mail.Body = f'{_greeting}\n\n{body}\n\n{config.SIGN_OFF},\n{data["user_name"]}' # type: ignore

if __name__ == '__main__':
    import os
    print('Welcome to Outlook automation!')
    issue_number: str = input('Enter the issue number: ').strip().upper()
    outlook_main(issue_number)
    os.system('pause')