Hello! I need your help to install a Python script. I am not a developer, so I need you to guide me through this process step-by-step. Please only give me one instruction at a time and wait for me to tell you I have completed it before moving on.

My goal is to set up the "Private iCal to Work Outlook Sync" script from this repository: https://github.com/mkorpela/calsync

Let's begin.

**Step 1: Prerequisites**

First, I need to make sure I have Python installed.

Tell me how to open a terminal on Windows (I can use either PowerShell or Command Prompt) and what command to type to check my Python version. Tell me what a successful output should look like.

*Wait for my response.*

**Step 2: Get the Code**

Next, I need to download the code.

Guide me to the GitHub repository page. Tell me to click the green "Code" button and then "Download ZIP". After I download it, instruct me to extract the ZIP file to a location I will remember, like `C:\calsync`.

*Wait for me to confirm I have extracted the files.*

**Step 3: Set up the Python Environment**

Now we need to set up a "virtual environment".

Tell me how to `cd` into the project directory I just created. Then, give me the two commands to create and activate the virtual environment. Tell me what my command prompt should look like after I activate it (it should have `(.venv)` at the beginning).

The commands are:
1. `python -m venv .venv`
2. `.\.venv\Scripts\activate`

*Wait for my response.*

**Step 4: Install Dependencies**

Now I need to install the required Python packages.

Give me the command `pip install -r requirements.txt`. Ask me to share the final lines of the output to confirm it was successful.

*Wait for my response.*

**Step 5: Register the Application in Azure**

This is the most complex part. Guide me carefully.

1.  Tell me to open a web browser and go to `portal.azure.com`, logging in with my **work** Microsoft 365 account.
2.  Once I'm in, instruct me to use the search bar at the top to find "Microsoft Entra ID" and click on it.
3.  Tell me to find "App registrations" in the left-hand menu and click it, then click "+ New registration".
4.  Now, walk me through filling out the registration form with the following details:
    *   **Name:** `Calendar Sync Utility`
    *   **Supported account types:** The default option (`Accounts in this organizational directory only`)
    *   **Redirect URI:** Tell me to select "Public client/native (mobile & desktop)" from the dropdown and enter `http://localhost` in the box.
5.  Tell me to click the "Register" button.
6.  After the app is created, tell me I will land on the "Overview" page. Instruct me to find and copy two values: the **Application (client) ID** and the **Directory (tenant) ID**. Tell me to save these in a temporary text file.
7.  Next, guide me to the "API permissions" page from the left-hand menu.
8.  Tell me to click "+ Add a permission", then select "Microsoft Graph", then "Delegated permissions".
9.  Instruct me to use the search box to find and check the boxes for two permissions: `Calendars.ReadWrite` and `User.Read`.
10. Tell me to click the "Add permissions" button at the bottom.

*Wait for me to confirm I have completed all these steps and have my two IDs saved.*

**Step 6: Configure the Script**

Now we will use the IDs I just saved.

1.  Tell me to go to my project folder (`C:\calsync`) and find the file `config.example.json`.
2.  Instruct me to make a copy of this file and rename the copy to `config.json`.
3.  Tell me to open `config.json` in a simple text editor like Notepad.
4.  Now, ask me for the following information one by one:
    *   The **Application (client) ID** I saved from Azure.
    *   The **Directory (tenant) ID** I saved from Azure.
    *   My private iCal URL. (Remind me how to find this in Google Calendar if I need it).
    *   My local timezone from this list: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones (e.g., "Europe/Helsinki").
5.  After I provide this information, show me exactly what my `config.json` file should look like with my details filled in. Also, show me an example of the `working_hours` so I can adjust it if I want to.

*Wait for me to confirm my `config.json` is saved.*

**Step 7: First Run**

We're ready to run the script for the first time.

1.  Tell me to go back to my terminal where the virtual environment is still active.
2.  Give me the command `pip install -e .` and tell me to run it.
3.  Once that is done, give me the final command to run the script: `python -m calsync_app.sync_calendar`.
4.  Warn me that a web browser will open, and I will need to log in with my work account and accept the permissions for the app we registered.
5.  Ask me to share the output from the terminal after the script finishes.

*Wait for my response.*

**Step 8: Automate the Script**

This is the final step. We will create a scheduled task to run the script automatically.

1.  Tell me to open PowerShell **as an Administrator**.
2.  Ask me for the full, absolute path to my project folder (e.g., `C:\calsync`).
3.  Once I provide the path, present me with the complete, filled-in PowerShell script below and instruct me to copy and paste the entire block into the administrative PowerShell window and press Enter.

```powershell
# --- 1. CONFIGURE THIS ---
$projectPath = "PASTE_THE_USER_PROVIDED_PATH_HERE"

# --- 2. DEFINE TASK DETAILS ---
$pythonExecutable = Join-Path $projectPath ".venv\Scripts\pythonw.exe"
$scriptArguments = "-m calsync_app.sync_calendar"
$taskAction = New-ScheduledTaskAction -Execute $pythonExecutable -Argument $scriptArguments -WorkingDirectory $projectPath
$taskTrigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 1) -Once -At (Get-Date)
$taskPrincipal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem).Username -LogonType Interactive
$taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

# --- 3. REGISTER THE TASK ---
Register-ScheduledTask -TaskName "Python Calendar Sync" -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal -Settings $taskSettings -Description "Syncs private iCal calendar to work Outlook calendar." -Force
Write-Host "Scheduled task 'Python Calendar Sync' has been created or updated successfully."
```

4.  Finally, tell me how to verify the task in the Windows Task Scheduler.

You have now guided me through the entire process. Thank you!
