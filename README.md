# 📧 Outlook Automation Tool

A lightweight desktop tool written in **Python** to automate repetitive email tasks using **Microsoft Outlook**.

This app allows users to:
- Automatically fill in recipient details
- Set email subject and body dynamically
- Include a greeting based on time of day (Good morning / Good evening)
- Sign off with a personalized signature
- Use localized messages via configuration

> The final app is compiled to `.exe` using **PyInstaller**, so Python is not required on the end user's system.
> You can download the standalone executable for your platform from the [Release Page](https://github.com/MariosMoraitis/Outlook_Automation/releases), and run it without installing Python.

---

## 🚀 Features

- 🕓 Time-based greeting (morning/evening)
- 📎 Custom signature
- 🌐 Multi-language support (Greek/English)
- ⚙️ User-configurable JSON files
- 🧠 Error handling with informative messages
- ✅ Works offline, locally

---

## 🧰 Requirements

- Windows 10/11
- Microsoft Outlook installed

For development only:
- Python 3.x
- `pywin32` package

---

## ⚙️ Configuration

### settings.json
This file defines global app behavior. Users can freely edit values, but should not change the keys.

```json
{
  "lang": "en",
  "signature": "Your Name"
}
```
lang: Language of greetings and signature block ("en" for English, "gr" for Greek)
signature: Will appear at the bottom of the email

### mail_config.json
Defines email content such as recipients and message body.
```json
{
  "recipients": "email1@example.com;email2@example.com",
  "cc_recipients": "cc1@example.com",
  "user_name": "Your Name",
  "subject": "",
  "body": "Please review the issue number "
}
```
recipients: Semicolon-separated primary recipients

cc_recipients: Semicolon-separated CC recipients

subject: If left blank, User's input will be inserted.

user_name: Will appear as part of the signature

body: The main part of the email — the issue number will be appended automatically.

## ▶️ How to Use
Run the compiled .exe file (no need for Python).

When prompted, enter the issue number.

Microsoft Outlook will open with a new email draft:

Recipients filled in

Subject set as the issue number

Body formatted with a greeting, message, and signature

You can review and send the email manually.
