import re
import mimetypes
from bs4 import BeautifulSoup
from libratom.lib.pff import PffArchive

def extract_header_field(headers, field_name):
    headers = headers or ""
    pattern = rf"^{field_name}:\s*(.*)$"
    match = re.search(pattern, headers, re.MULTILINE | re.IGNORECASE)
    return match.group(1).strip() if match else ""

def safe_getattr(obj, attr, default=""):
    try:
        return getattr(obj, attr, default)
    except Exception:
        return default

def get_attachment_info(attachment):
    try:
        name = attachment.name or "Unnamed"
        ext = name.split('.')[-1].lower() if '.' in name else ''
        mime_type, _ = mimetypes.guess_type(name)
        mime_type = mime_type or ext

        size = attachment.get_size()
        if size > 0:
            data = attachment.read_buffer(size)
        else:
            data = b""

        return {
            "name": name,
            "size": len(data),
            "type": mime_type,
            "data": data
        }
    except Exception as e:
        print(f"Attachment read error: {e}")
        return None

def extract_emails(pst_file):
    emails = []
    with PffArchive(pst_file) as archive:
        for message in archive.messages():
            subject = safe_getattr(message, "subject", "")
            sender = safe_getattr(message, "sender_name", "")
            sender_email = safe_getattr(message, "sender_email_address", "")
            sent_time = safe_getattr(message, "client_submit_time", "")
            headers = safe_getattr(message, "transport_headers", "")
            to = extract_header_field(headers, "To")
            cc = extract_header_field(headers, "Cc")

            plain_body = safe_getattr(message, "plain_text_body", None)
            html_body = safe_getattr(message, "html_body", None)

            if plain_body:
                body = plain_body
            elif html_body:
                soup = BeautifulSoup(html_body, "html.parser")
                body = soup.get_text()
            else:
                body = ""

            attachments = []
            try:
                for attachment in message.attachments:
                    att_info = get_attachment_info(attachment)
                    if att_info:
                        attachments.append(att_info)
            except Exception as e:
                print(f"Attachment list error: {e}")
                attachments = []

            emails.append({
                "subject": subject,
                "sender": f"{sender} <{sender_email}>",
                "to": to,
                "cc": cc,
                "sent_time": sent_time,
                "body": body,
                "attachments": attachments
            })

    return emails
    print(emails)
