# Risk Management

AI-powered risk and task extraction from meetings and emails.

---

## Quick Start (5 minutes)

### Step 1: Install Required Software

**On Mac:**
1. Install Homebrew (if not installed): https://brew.sh
2. Open Terminal and run:
   ```bash
   brew install python node git
   ```

**On Windows:**
1. Download and install Python: https://python.org/downloads (check "Add to PATH")
2. Download and install Node.js: https://nodejs.org (LTS version)
3. Download and install Git: https://git-scm.com/downloads

---

### Step 2: Download This Software

Open Terminal (Mac) or Command Prompt (Windows) and run:

```bash
git clone https://github.com/reqspace/risk-management.git
cd risk-management
```

---

### Step 3: Get Your API Key

1. Go to https://console.anthropic.com
2. Sign up or log in
3. Go to API Keys → Create Key
4. Copy the key (starts with `sk-ant-...`)

---

### Step 4: Configure

```bash
cp .env.example .env
```

Open `.env` in any text editor and paste your API key:
```
ANTHROPIC_API_KEY=sk-ant-your-key-here
```

---

### Step 5: Run It

**Mac:** Double-click `start.command`

**Windows:** Double-click `start.bat`

Your browser will open to the dashboard automatically.

---

## Need Help Setting Up?

Install Claude Code (AI assistant that can help):

```bash
npm install -g @anthropic-ai/claude-code
```

Then in Terminal, navigate to this folder and run:

```bash
cd risk-management
claude
```

Then paste this into Claude:

```
Help me set up Risk Management. Walk me through configuring my .env file and getting everything running.
```

Claude will guide you through everything step by step.

---

## How It Works

1. **Drop transcripts** into your project's `Transcripts/` folder
2. **AI extracts** risks and tasks automatically
3. **View everything** in the web dashboard
4. **Get daily emails** with status updates

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Python not found" | Reinstall Python, check "Add to PATH" |
| "npm not found" | Reinstall Node.js |
| Dashboard won't load | Check that both Terminal windows stay open |
| API errors | Verify your API key in `.env` |

---

## Support

Open an issue: https://github.com/reqspace/risk-management/issues
