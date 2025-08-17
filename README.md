# Private iCal to Work Outlook Sync

This Python script syncs events from a private iCal calendar (like a Google Calendar) to a work Outlook 365 calendar. It creates private events in Outlook with the generic subject "Personal Commitment" and marks the time as "Busy", effectively blocking out your personal time without exposing any private details.

The script is stateful and robust:
- It handles creating, updating (time changes), and deleting events.
- It filters events to only sync those that are marked as "Busy".
- It can be configured to only sync events that fall within your defined "working hours".
- It handles duplicate events in the source calendar gracefully.
- It is safe to run automatically on a schedule.

The source code is available at: https://github.com/mkorpela/calsync

## Prerequisites

1.  **Python 3.10+**: Make sure Python is installed on your Windows machine.
2.  **Private iCal URL**: You need the "secret address in iCal format" for the private calendar you want to sync.
    -   *For Google Calendar*: Go to `Settings` > `Settings for my calendars` > select your calendar > `Integrate calendar`. Copy the "Secret address in iCal format".
3.  **Work Microsoft 365 Account**: You must have the ability to log into your work account via a web browser.

## Setup Instructions

**Not a developer?** For a step-by-step, interactive installation, we strongly recommend using our AI-powered guide. Simply copy the contents of the file below and paste it into a chat with an AI assistant like Google Gemini or ChatGPT.

➡️ **[AI Installation Guide](./AI_INSTALL_GUIDE.md)**

---

For users comfortable with a manual setup, follow the instructions below.

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
2.  Open `config.json` with a text editor and fill in the values:
    -   `"client_id"`: Paste the **Application (client) ID** you copied from Azure.
    -   `"authority"`: In the URL, replace `YOUR_TENANT_ID_HERE` with the **Directory (tenant) ID** you copied.
    -   `"calendar_url"`: Paste your private iCal URL.
    -   `"timezone"`: Specify your local timezone (e.g., `"America/New_York"`, `"Europe/London"`). This ensures the "working hours" filter works correctly. A full list of valid timezones can be found on [Wikipedia](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones).
    -   `"working_hours"`: Define the days and times you want the script to sync. Events outside these hours will be ignored.
        -   Days are lowercase: `"monday"`, `"tuesday"`, etc.
        -   Times are in 24-hour format (`"HH:MM"`).
        -   You can have multiple time slots for a single day. This is useful for split schedules (e.g., working in the evening).
        -   If a day is not listed, no events will be synced for that day.

    **Example `working_hours` configuration:**
    ```json
    "working_hours": {
      "monday":    [{"start": "09:00", "end": "17:00"}, {"start": "21:00", "end": "23:00"}],
      "tuesday":   [{"start": "09:00", "end": "17:00"}],
      "wednesday": [{"start": "09:00", "end": "17:00"}],
      "thursday":  [{"start": "09:00", "end": "17:00"}],
      "friday":    [{"start": "09:00", "end": "13:00"}]
    }
    ```

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

## Automating the Sync with Windows Task Scheduler

To make this a "set it and forget it" utility, you can create a scheduled task to run the script automatically every hour. This can be done easily with a single PowerShell command.

1.  Open **PowerShell as an Administrator**.
2.  Copy the entire script block below.
3.  **IMPORTANT**: You **must** change the placeholder path in the `$projectPath` variable to the absolute path of your project folder.
4.  Paste and run the script in the administrative PowerShell window.

```powershell
# --- 1. CONFIGURE THIS ---
# IMPORTANT: Change this to the full, absolute path of your project's root directory.
$projectPath = "C:\path\to\your\calsync\project"

# --- 2. DEFINE TASK DETAILS ---
# Path to the "windowless" Python executable in the virtual environment
$pythonExecutable = Join-Path $projectPath ".venv\Scripts\pythonw.exe"
# Arguments to run the script as a module
$scriptArguments = "-m calsync_app.sync_calendar"

# Define the command to run
$taskAction = New-ScheduledTaskAction -Execute $pythonExecutable -Argument $scriptArguments -WorkingDirectory $projectPath
# Define when to run (every hour)
$taskTrigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 1) -Once -At (Get-Date)
# Define the user to run as
$taskPrincipal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem).Username -LogonType Interactive
# Define task settings
$taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

# --- 3. REGISTER THE TASK ---
Register-ScheduledTask -TaskName "Python Calendar Sync" -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal -Settings $taskSettings -Description "Syncs private iCal calendar to work Outlook calendar." -Force

Write-Host "Scheduled task 'Python Calendar Sync' has been created or updated successfully."
```

### Verifying the Task
1.  Open the **Task Scheduler** app from the Start Menu.
2.  Click on **Task Scheduler Library** on the left.
3.  You should see **Python Calendar Sync** in the list.
4.  You can right-click it to run it manually or view its history. A "Last Run Result" of `0x0` indicates success.

## License

This project is licensed under the MIT License. See the [LICENSE](./LICENSE) file for details
