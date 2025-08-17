# Private iCal to Work Outlook Sync

This Python script syncs events from a private iCal calendar (like a Google Calendar) to a work Outlook 365 calendar. It creates private events in Outlook with the generic subject "Personal Commitment" and marks the time as "Busy", effectively blocking out your personal time without exposing any private details.

The script is stateful and robust:
- It handles creating, updating (time changes), and deleting events.
- It uses unique event IDs to avoid creating duplicates on subsequent runs.
- It is safe to run automatically on a schedule.

## Prerequisites

1.  **Python 3.10+**: Make sure Python is installed on your Windows machine.
2.  **Private iCal URL**: You need the "secret address in iCal format" for the private calendar you want to sync.
    -   *For Google Calendar*: Go to `Settings` > `Settings for my calendars` > select your calendar > `Integrate calendar`. Copy the "Secret address in iCal format".
3.  **Work Microsoft 365 Account**: You must have the ability to log into your work account via a web browser.

## Setup Instructions

### 1. Get the Code
Clone this repository to your machine or download the source code as a ZIP file.

### 2. Set up a Python Environment
It is highly recommended to use a virtual environment. Open a terminal (PowerShell or CMD) in the project directory and run:
```bash
python -m venv .venv
.\.venv\Scripts\activate
```

### 3. Install Dependencies
Install the necessary Python packages:
```bash
pip install -r requirements.txt
```

### 4. Register an Application in Azure
This is the most critical step. You must register the script as an application in your company's Microsoft Entra ID (Azure) to grant it permission to access your calendar.

1.  Navigate to `portal.azure.com` and log in with your work account.
2.  Search for and navigate to **Microsoft Entra ID**.
3.  Go to **App registrations** and click **+ New registration**.
4.  Fill out the form:
    -   **Name:** `Calendar Sync Utility` (or any name you prefer).
    -   **Supported account types:** Leave the default (`Accounts in this organizational directory only`).
    -   **Redirect URI:** Select **Public client/native (mobile & desktop)** and enter `http://localhost`.
5.  Click **Register**.
6.  On the app's **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID**. You will need these for the configuration file.
7.  Go to the **API permissions** page for your app.
    -   Click **+ Add a permission**.
    -   Select **Microsoft Graph**.
    -   Select **Delegated permissions**.
    -   Search for and add the following two permissions:
        -   `Calendars.ReadWrite`
        -   `User.Read`
    -   Click **Add permissions**.

### 5. Configure the Script
1.  Make a copy of the `config.example.json` file and rename it to `config.json`.
2.  Open `config.json` and fill in the values:
    -   `"client_id"`: Paste the **Application (client) ID** you copied from Azure.
    -   `"authority"`: In the URL, replace `YOUR_TENANT_ID_HERE` with the **Directory (tenant) ID** you copied.
    -   `"calendar_url"`: Paste your private iCal URL.

### 6. Install the Local Package
To ensure the script and its tests can be found by Python, install it in "editable" mode. This links your virtual environment to your source code.
```bash
pip install -e .
```

## Usage

### First Run
Run the script from your terminal:
```bash
python -m calsync_app.sync_calendar
```
The first time you run it, a web browser window will open, asking you to log in to your Microsoft account and consent to the permissions you requested. This is a one-time process.

### Subsequent Runs
Every subsequent run will use a cached token (stored in `.msal_token_cache.json`) and will not require you to log in again.
```bash
python -m calsync_app.sync_calendar
```

### Scheduling
You can use the Windows Task Scheduler to run this script automatically (e.g., every hour). Create a new task that runs a command like:
`C:\path\to\your\project\.venv\Scripts\python.exe -m calsync_app.sync_calendar`

---