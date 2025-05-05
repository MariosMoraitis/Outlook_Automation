# 📧 Outlook Automation

**Outlook Automation** is a Python-based app that helps you quickly prepare Outlook email drafts based on predefined templates, greetings, and user settings.

Now featuring a **simple and modern GUI** built with [Eel](https://github.com/ChrisKnott/Eel) (Python + HTML/JS)!

Download from: https://github.com/MariosMoraitis/Outlook_Automation/releases/tag/version_2.1

---

## 🚀 Features

- 📧 Automatically generate a draft email in Outlook.
- 🌐 Supporting languages: English & Greek.
- ✍️ Optionally include your Outlook signature.
- 🛠 Easily configure settings from a web-based GUI.
- 🖥️ Standalone executable (`.exe`) version available.

---

## 🛠 How to Use

### 1. Run the Application

- If you're running from source:
  ```bash
  python main_gui.py
  ```
- If you're using the packaged .exe:

  Just double-click main_gui.exe.

➡️ A small browser window will open with the GUI.

### 2. Send an Email
  Enter the issue number in the input box.

  Click "Send Mail".

  An Outlook draft will open automatically, pre-filled according to your settings.

### 3. Edit Settings
  Click the "Settings" button on the main page.

  You can edit:

  Language: gr for Greek or en for English.

  Signature: yes to include Outlook signature, no to omit it.

  Press Save to update settings.json.

  After saving, you'll automatically return to the main page.

  ⚡ Note:
  New settings are applied instantly without needing to restart the app!

## ⚙️ Building the Executable (Optional)
If you want to build your own .exe:
```bash
pyinstaller --noconsole --onefile --add-data "web;web" --add-data "parameters;parameters" main_gui.py
```
✅ This bundles everything, including your HTML/CSS/JS and settings.

## 📋 Requirements
Python 3.10+
Eel
pywin32 (for Outlook automation)
Outlook installed and configured on your PC

Install dependencies:
```bash
pip install -r requirements.txt
```

## 📬 Contact
If you have any questions, suggestions, or just want to connect, feel free to reach out:

Name: Marios Moraitis

Email: mariosmorait.01@gmail.com

GitHub: https://github.com/MariosMoraitis

LinkedIn: https://www.linkedin.com/in/marios-moraitis-510539237/

Feel free to open issues or contribute!!!
