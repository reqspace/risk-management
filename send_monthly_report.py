"""
Send Monthly Risk Report with Executive Summary and Risk Register attachments.
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


def generate_monthly_report_html():
    """Generate monthly report with strong executive summary."""
    from daily_digest import get_portfolio_summary, get_critical_alerts, read_risk_register

    today = datetime.now().strftime('%B %d, %Y')
    month_year = datetime.now().strftime('%B %Y')

    portfolio = get_portfolio_summary()
    alerts = get_critical_alerts()

    # Count totals
    total_risks = sum(p['active_risks'] for p in portfolio)
    total_critical = len(alerts)
    critical_projects = len([p for p in portfolio if p['health'] == 'Critical'])

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
        .risk-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            font-size: 13px;
        }}
        .risk-table th {{
            background: #1e3a5f;
            color: white;
            padding: 10px;
            text-align: left;
        }}
        .risk-table td {{
            padding: 10px;
            border-bottom: 1px solid #dee2e6;
        }}
        .risk-table tr:nth-child(even) {{
            background: #f8f9fa;
        }}
        .high {{
            color: #dc3545;
            font-weight: bold;
        }}
        .recommendations {{
            background: #d4edda;
            border: 1px solid #28a745;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 0;
        }}
        .recommendations h3 {{
            color: #155724;
            margin-top: 0;
        }}
        .recommendations ul {{
            margin: 0;
            padding-left: 20px;
        }}
        .recommendations li {{
            margin: 8px 0;
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
        <h1>LDES Program Monthly Risk Report</h1>
        <div class="subtitle">{month_year} | Prepared {today}</div>
    </div>

    <div class="executive-summary">
        <h2>EXECUTIVE SUMMARY</h2>

        <p><strong>Portfolio Health: CRITICAL</strong> - {critical_projects} of 3 projects require immediate executive attention.</p>

        <div class="critical-alert">
            <h3>CRITICAL: RH PROJECT AT STANDSTILL</h3>
            <p><strong>Issue:</strong> EOS Energy has failed to deliver UL certification for battery containers per contractual agreement. This is a material breach blocking all permitting and construction activities.</p>
            <p><strong>Impact:</strong> Project timeline at risk. COD target of May 7, 2026 in jeopardy. Zero schedule float remaining.</p>
            <p><strong>Status:</strong> Awaiting EOS update on January 9, 2026. New ETA for UL cert is January 15, 2026.</p>
            <p><strong>Required Action:</strong> Executive engagement with EOS leadership. Evaluate contract remedies if deadline missed.</p>
        </div>

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
                <div class="value">3</div>
                <div class="label">Total Projects</div>
            </div>
        </div>

        <div class="recommendations">
            <h3>Immediate Actions Required</h3>
            <ul>
                <li><strong>RH - EOS UL Certification:</strong> Schedule executive call with EOS Energy CEO by January 6, 2026 to confirm January 15 delivery commitment</li>
                <li><strong>RH - Contingency Planning:</strong> Legal review of contract terms and remedies for continued non-performance</li>
                <li><strong>MCBCP - ATO Process:</strong> Initiate DoD Authorization to Operate process in Q1 2026 - long lead time item</li>
                <li><strong>CEC - ITC Timeline:</strong> Accelerate all project timelines to capture federal tax credits before potential policy changes</li>
            </ul>
        </div>
    </div>

    <div class="section">
        <h2>PROJECT STATUS DETAILS</h2>

        <div class="project-section critical">
            <h3>RH - 4MWh LDES Microgrid <span class="status-badge status-critical">CRITICAL</span></h3>
            <p><strong>Client:</strong> T2 Companies | <strong>EPC:</strong> Global Power Group | <strong>COD Target:</strong> May 7, 2026</p>
            <p><strong>Status:</strong> Project at standstill due to EOS UL Certification non-compliance. All permitting blocked pending UL listing.</p>

            <table class="risk-table">
                <tr>
                    <th>Risk ID</th>
                    <th>Risk</th>
                    <th>Probability</th>
                    <th>Impact</th>
                    <th>Owner</th>
                </tr>
                <tr>
                    <td>R-RH-001</td>
                    <td>EOS UL Certification - blocking permits</td>
                    <td class="high">High</td>
                    <td class="high">Critical</td>
                    <td>EOS Energy</td>
                </tr>
                <tr>
                    <td>R-RH-002</td>
                    <td>Permitting Delays - SD County timing</td>
                    <td class="high">High</td>
                    <td class="high">High</td>
                    <td>T2 Companies</td>
                </tr>
                <tr>
                    <td>R-RH-006</td>
                    <td>Schedule Float - zero float for completion</td>
                    <td class="high">High</td>
                    <td class="high">High</td>
                    <td>Global Power Group</td>
                </tr>
            </table>

            <p><strong>Key Milestones:</strong></p>
            <ul>
                <li>EOS UL Listing: <span class="high">DELAYED</span> (Baseline: Dec 1, 2025 | New ETA: Jan 15, 2026)</li>
                <li>Permitting Approval: <span class="high">AT RISK</span> (Target: Dec 15, 2025 - pending UL)</li>
                <li>Construction Start: <span class="high">AT RISK</span> (Target: Dec 19, 2025 - pending permits)</li>
                <li>IFC Design Set Approved: On Track (Jan 9, 2026)</li>
                <li>COD: On Track (May 7, 2026) - <em>if UL cert received by Jan 15</em></li>
            </ul>
        </div>

        <div class="project-section critical">
            <h3>MCBCP - 400 MWh Military Installation <span class="status-badge status-critical">CRITICAL</span></h3>
            <p><strong>Client:</strong> US Navy / DoD | <strong>Location:</strong> Camp Pendleton | <strong>Status:</strong> In Development</p>
            <p><strong>Status:</strong> 30% Design Set complete (Oct 2025). Critical path items: SDG&E interconnection and DoD ATO process.</p>

            <table class="risk-table">
                <tr>
                    <th>Risk ID</th>
                    <th>Risk</th>
                    <th>Probability</th>
                    <th>Impact</th>
                    <th>Owner</th>
                </tr>
                <tr>
                    <td>R-MCBCP-003</td>
                    <td>Cybersecurity/ATO - DoD Authorization required</td>
                    <td class="high">High</td>
                    <td class="high">Critical</td>
                    <td>T2 Companies</td>
                </tr>
                <tr>
                    <td>R-MCBCP-001</td>
                    <td>SDG&E Interconnection at Haybarn Canyon</td>
                    <td class="high">High</td>
                    <td class="high">High</td>
                    <td>Indian Energy</td>
                </tr>
            </table>
        </div>

        <div class="project-section">
            <h3>CEC - California LDES Program <span class="status-badge status-at-risk">AT RISK</span></h3>
            <p><strong>Goal:</strong> 30 GW PV + 450 GWh LDES by 2035 | <strong>Status:</strong> Policy work progressing</p>
            <p><strong>Status:</strong> NQC reform white paper submitted. Working on LDES Center of Excellence proposal and $50M BCP grant application.</p>

            <table class="risk-table">
                <tr>
                    <th>Risk ID</th>
                    <th>Risk</th>
                    <th>Probability</th>
                    <th>Impact</th>
                    <th>Owner</th>
                </tr>
                <tr>
                    <td>R-CEC-005</td>
                    <td>ITC Timeline - federal tax credit urgency</td>
                    <td class="high">High</td>
                    <td class="high">High</td>
                    <td>Industry</td>
                </tr>
                <tr>
                    <td>R-CEC-006</td>
                    <td>Interconnection Queue friction points</td>
                    <td class="high">High</td>
                    <td class="high">High</td>
                    <td>CAISO/CEC</td>
                </tr>
            </table>
        </div>
    </div>

    <div class="section">
        <h2>NEXT 30 DAYS - KEY ACTIONS</h2>
        <table class="risk-table">
            <tr>
                <th>Date</th>
                <th>Action</th>
                <th>Owner</th>
                <th>Project</th>
            </tr>
            <tr>
                <td>Jan 9, 2026</td>
                <td>EOS UL cert status update call</td>
                <td>T2 Companies</td>
                <td>RH</td>
            </tr>
            <tr>
                <td>Jan 9, 2026</td>
                <td>IFC Design Set Approval</td>
                <td>Global Power Group</td>
                <td>RH</td>
            </tr>
            <tr>
                <td>Jan 15, 2026</td>
                <td>EOS UL Certification ETA</td>
                <td>EOS Energy</td>
                <td>RH</td>
            </tr>
            <tr>
                <td>Jan 22-23, 2026</td>
                <td>Set 5 Battery Containers</td>
                <td>Global Power Group</td>
                <td>RH</td>
            </tr>
            <tr>
                <td>Jan 2026</td>
                <td>Draft Testing & Commissioning Plan</td>
                <td>T2 Companies</td>
                <td>RH</td>
            </tr>
            <tr>
                <td>Q1 2026</td>
                <td>Initiate ATO Process</td>
                <td>T2 Companies</td>
                <td>MCBCP</td>
            </tr>
            <tr>
                <td>Q1 2026</td>
                <td>$50M BCP Grant Application</td>
                <td>MFI</td>
                <td>CEC</td>
            </tr>
        </table>
    </div>

    <div class="footer">
        <p>LDES Program Monthly Risk Report | Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
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
    msg['Subject'] = f'LDES Program Monthly Risk Report - {month_year} - ACTION REQUIRED'
    msg['From'] = EMAIL_FROM
    msg['To'] = recipient

    # HTML content
    msg_alternative = MIMEMultipart('alternative')
    html_part = MIMEText(html_content, 'html')
    msg_alternative.attach(html_part)
    msg.attach(msg_alternative)

    # Attach Risk Register files
    risk_registers_path = BASE_PATH / 'Risk_Registers'

    for project in ['RH', 'MCBCP', 'CEC']:
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
    recipient = sys.argv[1] if len(sys.argv) > 1 else 'mlaporte@iepwr.com'
    send_monthly_report(recipient)
