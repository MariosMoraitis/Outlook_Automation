import eel
import json
from outlook_automation import outlook_main

eel.init('web')

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
        return f"❌ Error: {str(e)}"

@eel.expose
def send_email(issue_number):    
    try:
        outlook_main(issue_number)
    except Exception as e:
        return str(e)
    
    return "✅ Email prepared!"

eel.start('index.html', size=(500, 500))