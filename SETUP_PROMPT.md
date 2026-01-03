# Setup Prompt for Risk Management

Copy everything below this line and paste it into Claude Code:

---

Set up Risk Management on my system. This is an interactive setup - ask me questions and wait for my answers before proceeding.

## Step 1: Project Configuration
Ask me:
- What is my company/organization name?
- What are my project names/acronyms? (I can have 1 or more, e.g., "Project Alpha, PB, CEC")
- Do I have a logo file to use? If yes, ask me to place it in this folder and tell you the filename.

## Step 2: Create Project Structure
For each project name I provide:
- Create folder: `{project}/Transcripts/`
- Create folder: `{project}/Emails/`
- Create Excel file: `Risk_Registers/Risk_Register_{project}.xlsx` with these sheets:
  - **Risk Register** (columns: Risk ID, Title, Description, Category, Probability, Impact, Risk Score, Status, Trend, Owner, Mitigation Plan, Linked Tasks, Date Identified, Last Updated, Source)
  - **Tasks** (columns: Task ID, Task, Owner, Due Date, Status, Linked Risk, Source, Created Date, Completed Date)
  - **Update Log** (columns: Timestamp, Source, Source Type, Changes Made)
  - **Milestones** (columns: Milestone, Baseline Date, Current Date, Variance, Status)

## Step 3: Install Dependencies
- Install Python dependencies: `pip3 install flask anthropic openpyxl python-dotenv watchdog apscheduler requests python-docx`
- Install Node.js dependencies: `cd dashboard && npm install`

## Step 4: API Keys & Credentials
Ask me for each of these ONE AT A TIME:
1. Anthropic API key (tell me to get from console.anthropic.com if I don't have one)
2. ngrok auth token (tell me to get from dashboard.ngrok.com if I don't have one)
3. My email address (for sending reports)
4. My email app password (explain how to create one in Outlook/Gmail if needed)
5. Email recipients (who should receive the daily digest)

## Step 5: Configure the System
- Copy `.env.example` to `.env`
- Fill in all values in `.env` with my answers
- Update `server.py` to watch my specific project folders
- Update the dashboard config to show my project names
- Configure ngrok with my auth token

## Step 6: Test the Installation
- Start the Flask server and verify it runs
- Start ngrok and show me my public URL
- Start the dashboard
- Open localhost:3000 in my browser

## Step 7: Power Automate Setup
Give me step-by-step instructions with screenshot descriptions for:

**Outlook Email Automation:**
1. Create new Power Automate flow
2. Trigger: When email arrives in AI-Review folder
3. Action: HTTP POST to my ngrok URL with email content
4. Help me create Outlook folder structure: `AI-Review/{project}` for each of my projects

**Teams Meeting Transcripts (if I use Teams):**
1. Create new Power Automate flow
2. Trigger: When new file created in OneDrive Recordings folder
3. Action: HTTP POST to my ngrok URL with transcript content

## Step 8: Create Start Scripts

Create platform-specific scripts:

**start.bat (Windows):**
- Start Flask server in background
- Start ngrok tunnel
- Start dashboard
- Open browser to localhost:3000
- Display ngrok URL for Power Automate

**start.sh (Mac/Linux):**
- Same functionality as Windows version
- Make executable with `chmod +x`

## Step 9: Final Verification
- Confirm all services are running
- Show me the dashboard with my projects
- Show me my ngrok URL to use in Power Automate
- Give me a quick test: "Drop a test.txt file in one of your project Transcripts folders"

---

**Throughout this process:**
- Wait for my response before moving to the next step
- If something fails, help me troubleshoot before continuing
- Summarize what was configured at the end
