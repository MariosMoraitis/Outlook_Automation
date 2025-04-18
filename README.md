# Welcome to: Outlook Automation!
This .exe app, allows you to automate Outlook emails .

## What You Need

- A Windows 10 or 11 computer  
- Microsoft Outlook installed and set up  
- The file `` in the same folder as the app  
- The `Outlook_Automation.exe` file (no need to install Python)  

## Files You Should See

- `Outlook_Automation.exe` - The application  
- `` - Configuration file for email details  
- `user_guide.pdf` - This guide
- `outlook_automation.py` - Source code

## How to Use It

1. Double-click `Outlook_Automation.exe`  
2. You will be asked to enter an issue number (e.g., ABC123)  
   This will be used as the email subject.  
3. The program will:  
   - Open Microsoft Outlook  
   - Create a new email  
   - Fill in the recipients, subject, and message body  
   - Add a greeting and your name as a sign-off  
   - Leave the email open so you can review or edit it  
4. You can then click Send (or make edits first).  

## Setup the .json File

Here is an example of what it should look like:

```json
{
  "recipients": "someone@example.com",
  "cc_recipients": "manager@example.com",
  "body": "Please review and take the necessary action regarding the issue.",
  "user_name": "Your Name"
}

You can edit the json data as you like!!!
