import eel
import json
import sys
import os
from outlook_automation import outlook_main
from error_log import show_error, log_error

# Fix path when bundled with PyInstaller
def resource_path(relative_path) -> str:
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller stores data files in a temp folder
        base_path = sys.MEIPASS     #type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

web_dir: str = resource_path('web')
eel.init(web_dir)

@eel.expose
def get_settings():
    with open('parameters/settings.json', 'r', encoding='utf-8') as f:
        return json.load(f)
    
@eel.expose
def update_settings(lang, signature):
    try:
        with open('parameters/settings.json', 'r', encoding='utf-8') as f:
            settings = json.load(f)
        
        settings['lang'] = lang.strip()
        settings['signature'] = signature.strip()

        with open('parameters/settings.json', 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)

        return "✅ Settings saved successfully!"
    
    except Exception as e:
        error = f"❌ Error: {str(e)}"
        log_error(error)
        return error

@eel.expose
def send_email(issue_number):    
    try:
        outlook_main(issue_number)
        return "✅ Email prepared!"
    except SystemExit as e:
        return show_error()
    except Exception as e:
        return show_error()


eel.start('index.html', size=(500, 500))