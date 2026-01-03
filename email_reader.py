"""
Email Reader Module - Fetches and processes emails with attachments.
Supports Outlook/Office365 via IMAP or Microsoft Graph API.
"""

import os
import imaplib
import email
import requests
import base64
from email.header import decode_header
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List, Dict
import tempfile

from dotenv import load_dotenv

load_dotenv()

# Email settings
IMAP_SERVER = os.getenv('IMAP_SERVER', 'outlook.office365.com')
IMAP_PORT = int(os.getenv('IMAP_PORT', 993))
EMAIL_USER = os.getenv('MS_USER_EMAIL', os.getenv('EMAIL_FROM'))
EMAIL_PASSWORD = os.getenv('IMAP_PASSWORD', os.getenv('EMAIL_PASSWORD'))

# Microsoft Graph API settings
MS_CLIENT_ID = os.getenv('MS_CLIENT_ID')
MS_CLIENT_SECRET = os.getenv('MS_CLIENT_SECRET')
MS_TENANT_ID = os.getenv('MS_TENANT_ID')
MS_USER_EMAIL = os.getenv('MS_USER_EMAIL')

# Supported attachment types
SUPPORTED_EXTENSIONS = ['.txt', '.md', '.docx', '.pdf', '.doc']


def decode_mime_header(header):
    """Decode email header to string."""
    if header is None:
        return ""
    decoded_parts = decode_header(header)
    result = []
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            result.append(part.decode(encoding or 'utf-8', errors='ignore'))
        else:
            result.append(part)
    return ''.join(result)


def extract_text_from_attachment(file_path: Path, extension: str) -> str:
    """Extract text content from attachment based on file type."""
    try:
        if extension in ['.txt', '.md']:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()

        elif extension == '.docx':
            from docx import Document
            doc = Document(file_path)
            return '\n'.join([para.text for para in doc.paragraphs])

        elif extension == '.doc':
            # Try using textract or antiword if available
            try:
                import textract
                return textract.process(str(file_path)).decode('utf-8')
            except:
                return f"[Could not extract .doc file - install textract]"

        elif extension == '.pdf':
            try:
                import PyPDF2
                with open(file_path, 'rb') as f:
                    reader = PyPDF2.PdfReader(f)
                    text = []
                    for page in reader.pages:
                        text.append(page.extract_text() or '')
                    return '\n'.join(text)
            except:
                return f"[Could not extract PDF - install PyPDF2]"

        return ""
    except Exception as e:
        return f"[Error extracting {extension}: {e}]"


def get_graph_access_token() -> Optional[str]:
    """Get Microsoft Graph API access token using client credentials."""
    if not all([MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID]):
        return None

    try:
        token_url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
        token_data = {
            'grant_type': 'client_credentials',
            'client_id': MS_CLIENT_ID,
            'client_secret': MS_CLIENT_SECRET,
            'scope': 'https://graph.microsoft.com/.default'
        }
        response = requests.post(token_url, data=token_data)
        response.raise_for_status()
        return response.json()['access_token']
    except Exception as e:
        print(f"[Email Reader] Graph API token error: {e}")
        return None


def fetch_emails_via_graph(
    days_back: int = 7,
    subject_filter: Optional[str] = None,
    sender_filter: Optional[str] = None
) -> List[Dict]:
    """Fetch emails with attachments via Microsoft Graph API."""
    access_token = get_graph_access_token()
    if not access_token:
        print("[Email Reader] Could not get Graph API token")
        return []

    if not MS_USER_EMAIL:
        print("[Email Reader] MS_USER_EMAIL not configured")
        return []

    results = []
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    try:
        # Calculate date filter
        since_date = (datetime.now() - timedelta(days=days_back)).strftime('%Y-%m-%dT00:00:00Z')

        # Build filter query
        filter_parts = [f"receivedDateTime ge {since_date}", "hasAttachments eq true"]
        if subject_filter:
            filter_parts.append(f"contains(subject, '{subject_filter}')")

        filter_query = " and ".join(filter_parts)

        # Fetch messages
        messages_url = f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/messages"
        params = {
            '$filter': filter_query,
            '$select': 'id,subject,from,receivedDateTime,body,hasAttachments',
            '$orderby': 'receivedDateTime desc',
            '$top': 50
        }

        response = requests.get(messages_url, headers=headers, params=params)
        response.raise_for_status()
        messages = response.json().get('value', [])

        print(f"[Email Reader] Graph API found {len(messages)} emails with attachments")

        for msg in messages:
            # Apply sender filter if specified
            sender = msg.get('from', {}).get('emailAddress', {}).get('address', '')
            if sender_filter and sender_filter.lower() not in sender.lower():
                continue

            # Fetch attachments for this message
            attachments_url = f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/messages/{msg['id']}/attachments"
            attach_response = requests.get(attachments_url, headers=headers)
            attach_response.raise_for_status()
            attachments_data = attach_response.json().get('value', [])

            attachments = []
            for att in attachments_data:
                if att.get('@odata.type') == '#microsoft.graph.fileAttachment':
                    filename = att.get('name', '')
                    extension = Path(filename).suffix.lower()

                    if extension in SUPPORTED_EXTENSIONS:
                        # Decode base64 content and extract text
                        content_bytes = base64.b64decode(att.get('contentBytes', ''))

                        with tempfile.NamedTemporaryFile(delete=False, suffix=extension) as tmp:
                            tmp.write(content_bytes)
                            tmp_path = Path(tmp.name)

                        content = extract_text_from_attachment(tmp_path, extension)
                        tmp_path.unlink()

                        attachments.append({
                            'filename': filename,
                            'type': extension,
                            'content': content,
                            'size': len(content)
                        })
                        print(f"[Email Reader] Extracted: {filename} ({len(content)} chars)")

            if attachments:
                results.append({
                    'subject': msg.get('subject', ''),
                    'sender': sender,
                    'date': msg.get('receivedDateTime', ''),
                    'body': msg.get('body', {}).get('content', '')[:500],
                    'attachments': attachments
                })

        print(f"[Email Reader] Processed {len(results)} emails with supported attachments")
        return results

    except requests.exceptions.HTTPError as e:
        error_detail = e.response.json() if e.response else str(e)
        print(f"[Email Reader] Graph API error: {error_detail}")
        return []
    except Exception as e:
        print(f"[Email Reader] Graph API error: {e}")
        return []


def connect_imap() -> Optional[imaplib.IMAP4_SSL]:
    """Connect to email server via IMAP."""
    if not EMAIL_USER or not EMAIL_PASSWORD:
        print("[Email Reader] ERROR: Email credentials not configured")
        print("  Set IMAP_PASSWORD in .env (or use MS_USER_EMAIL + IMAP_PASSWORD)")
        return None

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASSWORD)
        print(f"[Email Reader] Connected to {IMAP_SERVER} as {EMAIL_USER}")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"[Email Reader] IMAP login failed: {e}")
        print("  For Outlook, you may need an App Password or OAuth token")
        return None
    except Exception as e:
        print(f"[Email Reader] Connection error: {e}")
        return None


def fetch_emails_with_attachments(
    folder: str = 'INBOX',
    days_back: int = 7,
    subject_filter: Optional[str] = None,
    sender_filter: Optional[str] = None
) -> List[Dict]:
    """
    Fetch emails with attachments from the specified folder.
    Tries Microsoft Graph API first, falls back to IMAP if unavailable.

    Args:
        folder: Email folder to search (default: INBOX)
        days_back: Number of days to look back
        subject_filter: Only emails containing this in subject
        sender_filter: Only emails from this sender

    Returns:
        List of dicts with email info and extracted attachment content
    """
    # Try Microsoft Graph API first (preferred for Office365)
    if all([MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID, MS_USER_EMAIL]):
        print("[Email Reader] Using Microsoft Graph API...")
        results = fetch_emails_via_graph(days_back, subject_filter, sender_filter)
        if results or not EMAIL_PASSWORD:  # Return results or if no IMAP fallback available
            return results
        print("[Email Reader] Graph API returned no results, trying IMAP...")

    # Fall back to IMAP
    mail = connect_imap()
    if not mail:
        return []

    results = []

    try:
        # Select folder
        status, _ = mail.select(folder)
        if status != 'OK':
            print(f"[Email Reader] Could not select folder: {folder}")
            return []

        # Build search criteria
        since_date = (datetime.now() - timedelta(days=days_back)).strftime('%d-%b-%Y')
        search_criteria = f'(SINCE "{since_date}")'

        # Search for emails
        status, message_ids = mail.search(None, search_criteria)
        if status != 'OK':
            print("[Email Reader] Search failed")
            return []

        ids = message_ids[0].split()
        print(f"[Email Reader] Found {len(ids)} emails in last {days_back} days")

        for msg_id in ids:
            try:
                status, msg_data = mail.fetch(msg_id, '(RFC822)')
                if status != 'OK':
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                # Get email metadata
                subject = decode_mime_header(msg['Subject'])
                sender = decode_mime_header(msg['From'])
                date_str = msg['Date']

                # Apply filters
                if subject_filter and subject_filter.lower() not in subject.lower():
                    continue
                if sender_filter and sender_filter.lower() not in sender.lower():
                    continue

                # Get email body
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        if content_type == 'text/plain':
                            payload = part.get_payload(decode=True)
                            if payload:
                                body = payload.decode('utf-8', errors='ignore')
                                break
                else:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        body = payload.decode('utf-8', errors='ignore')

                # Process attachments
                attachments = []
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue

                        filename = part.get_filename()
                        if filename:
                            filename = decode_mime_header(filename)
                            extension = Path(filename).suffix.lower()

                            if extension in SUPPORTED_EXTENSIONS:
                                # Save attachment temporarily and extract text
                                with tempfile.NamedTemporaryFile(delete=False, suffix=extension) as tmp:
                                    tmp.write(part.get_payload(decode=True))
                                    tmp_path = Path(tmp.name)

                                content = extract_text_from_attachment(tmp_path, extension)

                                attachments.append({
                                    'filename': filename,
                                    'type': extension,
                                    'content': content,
                                    'size': len(content)
                                })

                                # Clean up temp file
                                tmp_path.unlink()

                                print(f"[Email Reader] Extracted: {filename} ({len(content)} chars)")

                if attachments:  # Only include emails with supported attachments
                    results.append({
                        'subject': subject,
                        'sender': sender,
                        'date': date_str,
                        'body': body[:500],  # Truncate body
                        'attachments': attachments
                    })

            except Exception as e:
                print(f"[Email Reader] Error processing email: {e}")
                continue

        mail.logout()
        print(f"[Email Reader] Processed {len(results)} emails with attachments")
        return results

    except Exception as e:
        print(f"[Email Reader] Error: {e}")
        try:
            mail.logout()
        except:
            pass
        return []


def process_email_attachments(project_code: str = 'HB', days_back: int = 7, subject_filter: str = None):
    """
    Fetch emails and process attachments through the risk extraction pipeline.

    Args:
        project_code: Project code to associate risks/tasks with
        days_back: How many days back to search
        subject_filter: Optional subject line filter

    Returns:
        Dict with processing results
    """
    from process import process_content

    emails = fetch_emails_with_attachments(
        days_back=days_back,
        subject_filter=subject_filter
    )

    if not emails:
        return {'success': False, 'message': 'No emails with attachments found'}

    results = []

    for email_data in emails:
        for attachment in email_data['attachments']:
            if attachment['content'] and len(attachment['content']) > 50:
                source_name = f"Email: {email_data['subject'][:50]} - {attachment['filename']}"

                print(f"[Email Reader] Processing: {source_name}")

                try:
                    result = process_content(
                        content=attachment['content'],
                        project_code=project_code,
                        source_type='email_attachment',
                        source_name=source_name
                    )

                    results.append({
                        'email_subject': email_data['subject'],
                        'attachment': attachment['filename'],
                        'risks_found': result.get('risks_added', 0),
                        'tasks_found': result.get('tasks_added', 0)
                    })
                except Exception as e:
                    print(f"[Email Reader] Processing error: {e}")
                    results.append({
                        'email_subject': email_data['subject'],
                        'attachment': attachment['filename'],
                        'error': str(e)
                    })

    return {
        'success': True,
        'emails_processed': len(emails),
        'attachments_processed': len(results),
        'results': results
    }


if __name__ == '__main__':
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == '--process':
        # Process emails and extract risks
        project = sys.argv[2] if len(sys.argv) > 2 else 'HB'
        result = process_email_attachments(project_code=project)
        print(f"\nResults: {result}")
    else:
        # Just list emails with attachments
        emails = fetch_emails_with_attachments(days_back=7)
        print(f"\nFound {len(emails)} emails with supported attachments:")
        for e in emails:
            print(f"  - {e['subject']}")
            for a in e['attachments']:
                print(f"      Attachment: {a['filename']} ({a['size']} chars)")
