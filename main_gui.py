import eel
from outlook_automation import outlook_main

eel.init('web')

@eel.expose
def send_email(issue_number):    
    try:
        outlook_main(issue_number)
    except Exception as e:
        return str(e)
    
    return "✅ Email prepared!"

eel.start('index.html', size=(400, 300))