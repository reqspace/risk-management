"""
Send Monthly Risk Report with Executive Summary and Risk Register attachments.
Generates dynamic content from Risk Register data.
"""

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# Email settings
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
EMAIL_FROM = os.getenv('EMAIL_FROM')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')

BASE_PATH = Path(__file__).parent.resolve()


def get_all_projects():
    """Get list of all configured projects from Risk_Registers folder."""
    risk_registers_path = BASE_PATH / 'Risk_Registers'
    projects = []
    if risk_registers_path.exists():
        for f in risk_registers_path.glob('Risk_Register_*.xlsx'):
            project = f.stem.replace('Risk_Register_', '')
            projects.append(project)
    return projects


def generate_monthly_report_html():
    """Generate monthly report with dynamic content from Risk Registers."""
    from daily_digest import get_portfolio_summary, get_critical_alerts

    today = datetime.now().strftime('%B %d, %Y')
    month_year = datetime.now().strftime('%B %Y')

    portfolio = get_portfolio_summary()
    alerts = get_critical_alerts()

    # Count totals
    total_risks = sum(p.get('active_risks', 0) for p in portfolio)
    total_critical = len(alerts)
    critical_projects = len([p for p in portfolio if p.get('health') == 'Critical'])
    total_projects = len(portfolio)

    # Generate critical alerts HTML
    alerts_html = ""
    for alert in alerts[:5]:  # Top 5 critical alerts
        alerts_html += f"""
        <div class="critical-alert">
            <h3>{alert.get('risk_id', 'N/A')}: {alert.get('title', 'Critical Risk')}</h3>
            <p><strong>Project:</strong> {alert.get('project', 'N/A')}</p>
            <p>{alert.get('description', '')[:200]}...</p>
        </div>
        """

    # Generate project sections HTML
    projects_html = ""
    for p in portfolio:
        health = p.get('health', 'Unknown')
        status_class = 'status-critical' if health == 'Critical' else 'status-at-risk' if health == 'At Risk' else 'status-on-track'
        border_class = 'critical' if health in ['Critical', 'At Risk'] else ''

        projects_html += f"""
        <div class="project-section {border_class}">
            <h3>{p.get('code', 'Unknown')} <span class="status-badge {status_class}">{health.upper()}</span></h3>
            <p><strong>Active Risks:</strong> {p.get('active_risks', 0)} |
               <strong>Open Tasks:</strong> {p.get('open_tasks', 0)} |
               <strong>Overdue:</strong> {p.get('overdue_tasks', 0)}</p>
        </div>
        """

    html = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
        }}
        .header {{
            background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%);
            color: white;
            padding: 30px;
            border-radius: 8px 8px 0 0;
        }}
        .header h1 {{
            margin: 0;
            font-size: 28px;
        }}
        .header .subtitle {{
            opacity: 0.9;
            font-size: 16px;
            margin-top: 5px;
        }}
        .executive-summary {{
            background: #fff3cd;
            border: 2px solid #ffc107;
            border-radius: 0;
            padding: 25px;
            margin: 0;
        }}
        .executive-summary h2 {{
            color: #856404;
            margin-top: 0;
            font-size: 20px;
            border-bottom: 2px solid #ffc107;
            padding-bottom: 10px;
        }}
        .critical-alert {{
            background: #dc3545;
            color: white;
            padding: 20px;
            margin: 15px 0;
            border-radius: 6px;
        }}
        .critical-alert h3 {{
            margin: 0 0 10px 0;
            font-size: 16px;
        }}
        .critical-alert p {{
            margin: 5px 0;
            font-size: 14px;
        }}
        .key-metrics {{
            display: flex;
            gap: 15px;
            margin: 20px 0;
            flex-wrap: wrap;
        }}
        .metric {{
            background: white;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 15px 20px;
            min-width: 120px;
            text-align: center;
        }}
        .metric .value {{
            font-size: 32px;
            font-weight: bold;
            color: #1e3a5f;
        }}
        .metric .label {{
            font-size: 12px;
            color: #666;
            text-transform: uppercase;
        }}
        .metric.critical .value {{
            color: #dc3545;
        }}
        .section {{
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            padding: 25px;
            margin-top: 20px;
        }}
        .section h2 {{
            color: #1e3a5f;
            font-size: 18px;
            margin-top: 0;
            padding-bottom: 10px;
            border-bottom: 2px solid #2d5a87;
        }}
        .project-section {{
            background: white;
            border-radius: 8px;
            padding: 20px;
            margin: 15px 0;
            border-left: 4px solid #2d5a87;
        }}
        .project-section.critical {{
            border-left-color: #dc3545;
        }}
        .project-section h3 {{
            margin: 0 0 10px 0;
            color: #1e3a5f;
        }}
        .status-badge {{
            display: inline-block;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
        }}
        .status-critical {{
            background: #dc3545;
            color: white;
        }}
        .status-at-risk {{
            background: #ffc107;
            color: #333;
        }}
        .status-on-track {{
            background: #28a745;
            color: white;
        }}
        .footer {{
            text-align: center;
            color: #666;
            font-size: 12px;
            padding: 20px;
            border-top: 1px solid #dee2e6;
            margin-top: 30px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Monthly Risk Report</h1>
        <div class="subtitle">{month_year} | Prepared {today}</div>
    </div>

    <div class="executive-summary">
        <h2>EXECUTIVE SUMMARY</h2>

        <p><strong>Portfolio Health:</strong> {critical_projects} of {total_projects} projects require attention.</p>

        {alerts_html if alerts_html else '<p>No critical alerts at this time.</p>'}

        <div class="key-metrics">
            <div class="metric critical">
                <div class="value">{total_critical}</div>
                <div class="label">Critical Alerts</div>
            </div>
            <div class="metric">
                <div class="value">{total_risks}</div>
                <div class="label">Active Risks</div>
            </div>
            <div class="metric critical">
                <div class="value">{critical_projects}</div>
                <div class="label">Critical Projects</div>
            </div>
            <div class="metric">
                <div class="value">{total_projects}</div>
                <div class="label">Total Projects</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2>PROJECT STATUS</h2>
        {projects_html if projects_html else '<p>No projects configured.</p>'}
    </div>

    <div class="footer">
        <p>Monthly Risk Report | Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <p>Risk Register spreadsheets attached for detailed review</p>
    </div>
</body>
</html>
"""
    return html


def send_monthly_report(recipient):
    """Send monthly report with Risk Register attachments."""
    html_content = generate_monthly_report_html()
    today = datetime.now().strftime('%Y-%m-%d')
    month_year = datetime.now().strftime('%B %Y')

    msg = MIMEMultipart('mixed')
    msg['Subject'] = f'Monthly Risk Report - {month_year}'
    msg['From'] = EMAIL_FROM
    msg['To'] = recipient

    # HTML content
    msg_alternative = MIMEMultipart('alternative')
    html_part = MIMEText(html_content, 'html')
    msg_alternative.attach(html_part)
    msg.attach(msg_alternative)

    # Attach Risk Register files
    risk_registers_path = BASE_PATH / 'Risk_Registers'
    projects = get_all_projects()

    for project in projects:
        file_path = risk_registers_path / f'Risk_Register_{project}.xlsx'
        if file_path.exists():
            with open(file_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                attachment.add_header('Content-Disposition', 'attachment', filename=f'Risk_Register_{project}_{today}.xlsx')
                msg.attach(attachment)
                print(f"[Monthly Report] Attached: Risk_Register_{project}.xlsx")

    # Send email
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(EMAIL_FROM, [recipient], msg.as_string())

    print(f"[Monthly Report] Email sent successfully to {recipient}")
    return True


if __name__ == '__main__':
    import sys
    recipient = sys.argv[1] if len(sys.argv) > 1 else os.getenv('EMAIL_TO', '')
    if recipient:
        send_monthly_report(recipient)
    else:
        print("Error: No recipient specified. Set EMAIL_TO in .env or pass as argument.")
