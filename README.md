# Risk Management

Automated risk and task extraction from meetings and emails using AI.

## Features

- **Automatic Risk Extraction**: Drop meeting transcripts and get risks/tasks automatically identified
- **Email Processing**: Process email attachments for risks and action items
- **Real-time Dashboard**: Web-based dashboard showing all projects, risks, and tasks
- **Daily Digest**: Automated email summaries of project status
- **Monthly Reports**: Generate formatted reports with attachments
- **Power Automate Integration**: Connect to Outlook and Teams for seamless automation

## Requirements

- Python 3.10+
- Node.js 18+
- Claude Code CLI
- ngrok account (free tier works)
- Anthropic API key
- Microsoft 365 (for Power Automate integration)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/YOUR_USERNAME/risk-management.git
   cd risk-management
   ```

2. Open Claude Code in the project folder:
   ```bash
   claude
   ```

3. Copy the contents of `SETUP_PROMPT.md` and paste into Claude Code

4. Follow the interactive setup wizard

## What the Setup Will Ask For

- Your project names/acronyms (e.g., "Project Alpha, PB, CEC")
- Your company name and logo
- Anthropic API key
- ngrok auth token
- Email settings for reports

## Quick Start (After Setup)

**Windows:**
```bash
start.bat
```

**Mac/Linux:**
```bash
./start.sh
```

This will:
- Start the Flask API server
- Start ngrok tunnel
- Launch the dashboard
- Open your browser to the dashboard

## Project Structure

```
risk-management/
├── server.py           # Flask API server
├── process.py          # AI risk extraction engine
├── daily_digest.py     # Daily email digest
├── monthly_report.py   # Monthly report generation
├── email_reader.py     # Email attachment processing
├── dashboard/          # React dashboard
│   ├── src/
│   │   ├── pages/
│   │   │   ├── Portfolio.jsx
│   │   │   ├── Dashboard.jsx
│   │   │   ├── RiskRegister.jsx
│   │   │   └── Tasks.jsx
│   │   └── components/
│   └── public/
├── Risk_Registers/     # Excel files (created by setup)
└── {Project}/          # Project folders (created by setup)
    ├── Transcripts/
    └── Emails/
```

## Usage

### Processing Transcripts
Drop `.txt`, `.md`, or `.docx` files into any project's `Transcripts/` folder. The system will automatically:
1. Extract risks and tasks using AI
2. Add them to the Risk Register
3. Log the update

### Dashboard
Access at `http://localhost:3000`:
- **Portfolio**: Overview of all projects
- **Risks**: View and filter all risks
- **Tasks**: View and filter all tasks
- **Reports**: Generate and send reports

### API Endpoints
- `GET /health` - Health check
- `POST /process` - Process content for risks/tasks
- `GET /api/portfolio` - Get all projects summary
- `GET /api/risks` - Get all risks
- `GET /api/tasks` - Get all tasks
- `POST /api/reports/send` - Send report email

## Support

For assistance, contact the developer or open an issue on GitHub.

## License

MIT License
