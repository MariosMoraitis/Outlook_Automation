from datetime import datetime
import win32com.client
import json
import os

def prepare_body(body, user_name) -> str:
    time: int = datetime.now().hour
    if time < 12:
        greeting: str = "Καλημέρα,\n\n"
    else:
        greeting: str = "Καλησπέρα,\n\n"
    
    sign_off: str = f'\n\nΕυχαριστώ,\n{user_name}'

    return greeting + body + sign_off

def load_json():
    try:
        with open('....json', 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Check for valid keys...
        required_keys: list[str] = ["recipients", "cc_recipients", "body", "user_name"]
        for key in required_keys:
            if key not in data:
                raise KeyError(f"Missing key '{key}' in JSON file.")

        return data
    
    except KeyError as e:
        print(f"\n⚠️ ERROR: {e}")
    except FileNotFoundError:
        print("\n⚠️ ERROR: JSON file: '....json' not found!!!.")
    except Exception as e:
        print("\n⚠️ ERROR: Unexpected error occurred...")

def main() -> None:
    print('Welcome to Outlook automation!')
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
    main()
    os.system('pause')