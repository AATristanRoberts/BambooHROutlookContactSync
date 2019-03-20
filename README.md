# BambooHR / Outlook Contact Sync
_Sync contacts from the employee directory in BambooHR to Microsoft Outlook._

# Setup / Install

## 1. Install Python 3.7+
Download [Python 3.7 for Windows](https://www.python.org/downloads/release/python-372/) and install it. It is recommended to use the Windows x86-64 executable installer linked to from the **Files** section of the page.

## 2. Install pywin32
Download [pywin32](https://github.com/mhammond/pywin32/releases) and install it. If you installed the x86-64 version of Python above, you will want the `pywin32-224.win-amd64-py3.7.exe` version.

## 3. Install Python Libraries
Open a Command Prompt and run:
```
python -m pip install requests pyppeteer
```

## 4. Install Google Chrome
If you do not already have it, download and install [Google Chrome](https://www.google.com/chrome/).

## 5. Download and Configure the `sync.py` script
Download the `sync.py` file from this repo and configure it by opening it in a text editor (e.g. Notepad / Notepad++ / Sublime Text) and editing two lines:

### Line 12: DIRECTORY URL
```
DIRECTORY_URL = "https://automationanywhere.bamboohr.com/employees/directory.php"
```

Unless you work for Automation Anywhere, change this to the URL of your employee directory.

### Line 14: FILTER_CARDS
```
FILTER_CARDS = "AA AU"  # A filter to apply to the contacts
```

By default, the program will filter the directory. Replace `AA AU` with the filter you want to apply. If you want to sync all employees, make it an empty string (i.e. `FILTER_CARDS = ""`).


# Running the Program

1. Close Microsoft Outlook if it is open. If you do not close Outlook, a `Server Execution Failed` error will occur.

2. Open a command prompt and navigate to the path you downloaded the `sync.py` script, e.g.: `cd %USERPROFILE%\Downloads`

3. Run the script: `python sync.py`

4. When the BambooHR page opens in Google Chrome, enter your login details and click Login.

5. Wait for the script to say `Done` then your contacts will be synced!

# Configuration
