import re
from libratom.lib.pff import PffArchive

def extract_header_field(headers, field_name):
    headers = headers or ""
    pattern = rf"^{field_name}:\s*(.*)$"
    match = re.search(pattern, headers, re.MULTILINE | re.IGNORECASE)
    return match.group(1).strip() if match else f"(No {field_name})"

def safe_getattr(obj, attr, default="(Unavailable)"):
    try:
        return getattr(obj, attr, default)
    except Exception as e:
        return f"{default} - Error: {e}"

def read_pst_with_libratom(pst_file_path):
    with PffArchive(pst_file_path) as archive:
        for message in archive.messages():
            print("=" * 60)

            subject = safe_getattr(message, "subject", "(No Subject)")
            sender = safe_getattr(message, "sender_name", "(Unknown Sender)")
            sender_email = safe_getattr(message, "sender_email_address", "(No Email)")
            headers = safe_getattr(message, "transport_headers", "")

            to = extract_header_field(headers, "To")
            cc = extract_header_field(headers, "Cc")
            sent_time = safe_getattr(message, "client_submit_time", "(No Date)")
            body = safe_getattr(message, "plain_text_body", "(No Body)") or safe_getattr(message, "html_body", "(No Body)")

            print(f"Subject: {subject}")
            print(f"Sender: {sender} <{sender_email}>")
            print(f"To: <{to}>")
            print(f"CC: {cc}")
            print(f"Sent on: {sent_time}")
            print(f"Body:\n{body}\n")

# Example usage
pst_path = r"C:\Users\Pranj\OneDrive\Desktop\Outlook\backup.pst"  # Replace with actual PST file path
read_pst_with_libratom(pst_path)
