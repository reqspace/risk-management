# Risk Management - Installation Guide

## Step 1: Download the App

1. Go to: https://github.com/reqspace/risk-management/releases
2. Download:
   - **Mac (Apple Silicon)**: `Risk.Management-X.X.X-arm64.dmg`
   - **Mac (Intel)**: `Risk.Management-X.X.X.dmg`
   - **Windows**: `Risk.Management.Setup.X.X.X.exe`

3. Install like any normal app:
   - **Mac**: Open DMG → Drag to Applications
   - **Windows**: Run installer → Follow prompts

---

## Step 2: First Launch

1. **Open the app** (Risk Management)
2. **Welcome dialog** appears - click "Open Config File"
3. **Add your API key** to the `.env` file that opens:
   ```
   ANTHROPIC_API_KEY=sk-ant-your-key-here
   ```
4. **Save the file** and close it
5. **Restart the app**

### Getting Your Anthropic API Key

1. Go to https://console.anthropic.com
2. Sign up or log in
3. Click "API Keys" → "Create Key"
4. Copy the key (starts with `sk-ant-...`)

---

## Step 3: Create Your First Project

1. In the app, the dashboard shows "No projects"
2. Open Finder/Explorer and go to:
   - **Mac**: `~/Library/Application Support/Risk Management/`
   - **Windows**: `%APPDATA%/Risk Management/`
3. Create folders:
   ```
   MyProject/
   ├── Transcripts/    ← Drop meeting notes here
   └── Emails/         ← Drop email exports here
   ```
4. Create a Risk Register Excel file:
   ```
   Risk_Registers/
   └── Risk_Register_MyProject.xlsx
   ```
5. **Restart the app** - your project appears!

---

## Step 4: Set Up Email Reports (Optional)

Add these to your `.env` file:

```
# Email settings for sending reports
SMTP_SERVER=smtp.office365.com
SMTP_PORT=587
EMAIL_FROM=your.email@company.com
EMAIL_PASSWORD=your-app-password
EMAIL_TO=recipient@company.com
```

### Getting an App Password (Microsoft 365)

1. Go to https://account.microsoft.com/security
2. Click "Advanced security options"
3. Under "App passwords", click "Create a new app password"
4. Copy the generated password to `EMAIL_PASSWORD`

---

## Step 5: Power Automate Integration (Optional)

This lets Outlook automatically send emails to the app for processing.

### Prerequisites
- Microsoft 365 account
- Power Automate access (included with most M365 plans)
- ngrok account (free at https://ngrok.com)

### Set Up ngrok

1. Sign up at https://ngrok.com
2. Get your auth token from the dashboard
3. Add to `.env`:
   ```
   NGROK_AUTH_TOKEN=your-ngrok-token
   ```
4. The app will show you your public URL when running

### Create Power Automate Flow

1. Go to https://make.powerautomate.com
2. Click "Create" → "Automated cloud flow"
3. Name: "Send to Risk Management"
4. Trigger: "When a new email arrives (V3)"
   - Folder: `Inbox/AI-Review` (create this folder in Outlook)
   - Include Attachments: Yes

5. Add action: "HTTP"
   - Method: POST
   - URI: `https://YOUR-NGROK-URL/api/process`
   - Headers: `Content-Type: application/json`
   - Body:
     ```json
     {
       "content": "@{triggerOutputs()?['body/body']}",
       "source": "@{triggerOutputs()?['body/subject']}",
       "project": "MyProject"
     }
     ```

6. Save and test by moving an email to `AI-Review` folder

---

## Step 6: Microsoft Graph Integration (Optional)

This allows the app to read your calendar for meeting-related risks.

### Create Azure App Registration

1. Go to https://portal.azure.com
2. Search for "App registrations" → "New registration"
3. Name: "Risk Management"
4. Supported account types: "Single tenant"
5. Click "Register"

6. Copy these values to `.env`:
   ```
   MS_CLIENT_ID=<Application (client) ID>
   MS_TENANT_ID=<Directory (tenant) ID>
   ```

7. Go to "Certificates & secrets" → "New client secret"
8. Copy the secret value to `.env`:
   ```
   MS_CLIENT_SECRET=<your-secret>
   ```

9. Go to "API permissions" → "Add a permission"
   - Microsoft Graph → Application permissions
   - Add: `Calendars.Read`, `Mail.Read`
   - Click "Grant admin consent"

10. Add your email to `.env`:
    ```
    MS_USER_EMAIL=your.email@company.com
    ```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| App won't open (Mac) | Right-click → Open → Open |
| "API key not set" | Check `.env` file has `ANTHROPIC_API_KEY=sk-ant-...` |
| Projects don't appear | Create folders in app data directory, restart app |
| Email reports fail | Check SMTP settings, try app password not regular password |
| Power Automate fails | Check ngrok URL is current (changes each restart) |

---

## File Locations

| Item | Mac | Windows |
|------|-----|---------|
| Config | `~/Library/Application Support/Risk Management/.env` | `%APPDATA%/Risk Management/.env` |
| Data | `~/Library/Application Support/Risk Management/` | `%APPDATA%/Risk Management/` |
| Logs | `~/Library/Logs/Risk Management/` | `%APPDATA%/Risk Management/logs/` |

---

## Getting Updates

The app checks for updates automatically. When an update is available:
1. A dialog appears with "Update Available"
2. Download happens in background
3. When complete, click "Restart" to install

---

## Support

- GitHub Issues: https://github.com/reqspace/risk-management/issues
- Documentation: https://github.com/reqspace/risk-management#readme
