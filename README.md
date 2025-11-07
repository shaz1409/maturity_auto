# Marketing Maturity Assessment Automation

Automated system to generate personalized marketing maturity assessment presentations from Google Sheets survey data.

## Features

- 📊 Loads survey data directly from Google Sheets
- 🎯 Calculates category scores for each client
- 🤖 Generates AI-powered recommendations using OpenAI
- 📑 Creates personalized PowerPoint presentations
- 🎨 Automatically positions score indicators on slides

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Set Environment Variables

#### OpenAI API Key (Required)

```bash
export OPENAI_API_KEY='your-api-key-here'
```

#### SharePoint Authentication (Optional - for automatic upload)

**Option A: App-Based Authentication (Recommended for Automation)**

This is the most reliable method for automated scripts. One-time setup:

1. Register an app in Azure AD:
   - Go to https://portal.azure.com
   - Navigate to **Azure Active Directory** → **App registrations**
   - Click **New registration**
   - Name: `Maturity Assessment Automation`
   - Supported account types: **Accounts in this organizational directory only**
   - Click **Register**

2. Create a client secret:
   - In your app, go to **Certificates & secrets**
   - Click **New client secret**
   - Description: `Maturity Assessment Script`
   - Expires: **24 months** (or your preference)
   - Click **Add** and **copy the secret value** (you won't see it again!)

3. Grant SharePoint permissions:
   - Go to **API permissions** → **Add a permission**
   - Select **Microsoft Graph** → **Application permissions** (⚠️ NOT Delegated)
   - Search for: `Sites.ReadWrite.All`
   - Check the box and click **Add permissions**
   - Click **Grant admin consent** (if you have permission)
   - ⚠️ **Important**: Application permissions are required for app-based auth

4. Set environment variables:
   ```bash
   export SHAREPOINT_UPLOAD=true
   export SHAREPOINT_AUTH_METHOD=app
   export SHAREPOINT_CLIENT_ID='your-client-id-from-azure'
   export SHAREPOINT_CLIENT_SECRET='your-client-secret-value'
   ```

**Option B: App Password (Quick Fix - if MFA is enabled)**

If your account has MFA enabled, you can use an app password:

1. Create an app password:
   - Go to https://account.microsoft.com/security
   - Under **App passwords**, create a new one
   - Copy the generated password

2. Set environment variables:
   ```bash
   export SHAREPOINT_UPLOAD=true
   export SHAREPOINT_AUTH_METHOD=user
   export SHAREPOINT_USERNAME='your-email@domain.com'
   export SHAREPOINT_PASSWORD='your-app-password-here'
   ```

**Make Environment Variables Persistent**

Add to your `~/.zshrc` (or `~/.bashrc`) for persistence:

```bash
# OpenAI
echo "export OPENAI_API_KEY='your-api-key-here'" >> ~/.zshrc

# SharePoint (App-based - recommended)
echo "export SHAREPOINT_UPLOAD=true" >> ~/.zshrc
echo "export SHAREPOINT_AUTH_METHOD=app" >> ~/.zshrc
echo "export SHAREPOINT_CLIENT_ID='your-client-id'" >> ~/.zshrc
echo "export SHAREPOINT_CLIENT_SECRET='your-client-secret'" >> ~/.zshrc

source ~/.zshrc
```

### 3. Configure Google Sheet

Update the `SHEET_ID` and `GID` in `maturity_assessment.py` if using a different Google Sheet.

## Usage

### Run the Script

```bash
python maturity_assessment.py
```

Or make it executable and run directly:

```bash
chmod +x maturity_assessment.py
./maturity_assessment.py
```

### Output

Generated PowerPoint presentations will be saved in the `output/` directory with filenames like:
- `client_email_at_domain_com_Maturity_Assessment.pptx`

## Scheduling

### macOS (using launchd)

1. Create a plist file at `~/Library/LaunchAgents/com.maturity.assessment.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.maturity.assessment</string>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/bin/python3</string>
        <string>/Users/shazahmed/Documents/python_repos/maturity_auto/maturity_assessment.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>/Users/shazahmed/Documents/python_repos/maturity_auto</string>
    <key>EnvironmentVariables</key>
    <dict>
        <key>OPENAI_API_KEY</key>
        <string>your-api-key-here</string>
        <key>SHAREPOINT_UPLOAD</key>
        <string>true</string>
        <key>SHAREPOINT_AUTH_METHOD</key>
        <string>app</string>
        <key>SHAREPOINT_CLIENT_ID</key>
        <string>your-client-id</string>
        <key>SHAREPOINT_CLIENT_SECRET</key>
        <string>your-client-secret</string>
    </dict>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>9</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>/Users/shazahmed/Documents/python_repos/maturity_auto/logs/output.log</string>
    <key>StandardErrorPath</key>
    <string>/Users/shazahmed/Documents/python_repos/maturity_auto/logs/error.log</string>
</dict>
</plist>
```

2. Load the job:

```bash
launchctl load ~/Library/LaunchAgents/com.maturity.assessment.plist
```

3. Check status:

```bash
launchctl list | grep maturity
```

### Linux (using cron)

Add to crontab (`crontab -e`):

```bash
# Run daily at 9:00 AM
0 9 * * * cd /path/to/maturity_auto && /usr/bin/python3 maturity_assessment.py >> logs/cron.log 2>&1
```

### Windows (using Task Scheduler)

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (daily at 9:00 AM)
4. Set action: Start a program
   - Program: `python`
   - Arguments: `maturity_assessment.py`
   - Start in: `C:\path\to\maturity_auto`

## Project Structure

```
maturity_auto/
├── maturity_assessment.py      # Main automation script
├── maturity_analysis.ipynb      # Jupyter notebook (for development)
├── Maturity_Slide_Template.pptx # PowerPoint template
├── requirements.txt             # Python dependencies
├── output/                      # Generated presentations (gitignored)
└── README.md                    # This file
```

## Categories

The assessment covers 5 categories:

1. **Tech & Data** (5 questions)
2. **Campaigning & Assets** (6 questions)
3. **Segmentation & Personalisation** (3 questions)
4. **Reporting & Insights** (6 questions)
5. **People & Operations** (4 questions)

## Requirements

- Python 3.8+
- OpenAI API key
- Google Sheet with survey responses
- PowerPoint template file

## License

[Your License Here]

