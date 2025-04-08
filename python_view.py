import os
import html
import uuid
import streamlit as st
from pst_utils import extract_emails

st.set_page_config(page_title="📁 Email Explorer", layout="wide")

def format_body(text):
    if isinstance(text, bytes):
        text = text.decode("utf-8", errors="ignore")
    return text.replace('\r\n', '\n').replace('\r', '\n').replace('\t', '    ').strip()

st.markdown("""
    <style>
    .main, .stApp {
        background-color: #0f1117;
        color: #e0e0e0;
    }
    .email-card {
        background-color: #1e212b;
        padding: 1.2rem;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.5);
        margin-bottom: 1.2rem;
        border-left: 6px solid #4e8cff;
    }
    .email-subject {
        font-weight: bold;
        font-size: 18px;
        color: #4e8cff;
    }
    .email-meta {
        font-size: 14px;
        color: #c0c0c0;
        margin-bottom: 0.8rem;
    }
    .email-body {
        font-size: 15px;
        color: #e0e0e0;
        white-space: pre-wrap;
    }
    input, textarea {
        background-color: #1e1e1e !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📂 PST Email Viewer")
st.markdown("🔹 Enter the full path to your `.pst` file on disk")

pst_path = st.text_input("📄 Full PST File Path", placeholder="e.g., D:\\Backups\\Outlook\\archive2023.pst")

if pst_path:
    if os.path.exists(pst_path):
        st.info(f"📂 Using PST file: `{pst_path}`")

        with st.spinner("⏳ Processing PST... this may take a while for large files..."):
            emails = extract_emails(pst_path)

        st.success(f"✅ Extracted {len(emails)} emails from your PST file!")

        query = st.text_input("🔍 Search by subject or body keyword")

        if query:
            query = query.lower()
            results = [
                email for email in emails
                if query in str(email.get("subject", "")).lower() or
                   query in str(email.get("body", "")).lower()
            ]
            st.markdown(f"### 🔎 Found **{len(results)}** emails containing '**{html.escape(query)}**'")

            for i, email in enumerate(results):
                subject = html.escape(str(email.get("subject", "(No Subject)")))
                sender = html.escape(str(email.get("sender", "(Unknown Sender)")))
                to = html.escape(str(email.get("to", "(No Recipient)")))
                cc = html.escape(str(email.get("cc", "")))
                sent_time = html.escape(str(email.get("sent_time", "(No Date)")))
                body_text = format_body(email.get("body", "(No Body)"))
                trimmed_body = html.escape(body_text[:3000])

                with st.container():
                    st.markdown(f"""
                        <div class="email-card">
                            <div class="email-subject">{subject}</div>
                            <div class="email-meta">
                                <strong>From:</strong> {sender}<br>
                                <strong>To:</strong> {to}<br>
                                <strong>CC:</strong> {cc}<br>
                                <strong>Sent:</strong> {sent_time}
                            </div>
                            <div class="email-body">{trimmed_body}</div>
                        </div>
                    """, unsafe_allow_html=True)

                    with st.expander("📋 Copy Full Email Body"):
                        st.code(body_text, language="text")

                    attachments = email.get("attachments", [])
                    if attachments:
                        with st.expander(f"📎 Attachments ({len(attachments)})"):
                            for att in attachments:
                                name = att.get("name", "Unnamed file")
                                size = att.get("size", 0)
                                content = att.get("data", b"")
                                size_kb = round(size / 1024, 2)

                                st.markdown(f"**📄 {name}** ({size_kb} KB)")
                                st.download_button(
                                    label="⬇️ Download",
                                    data=content,
                                    file_name=name,
                                    mime="application/octet-stream",
                                    key=f"download-{i}-{name}-{uuid.uuid4()}"
                                )
                    else:
                        st.markdown("*No attachments.*")
        else:
            st.info("Enter a keyword above to search your emails.")
    else:
        st.error("🚫 File path not found. Please enter a valid `.pst` file path.")
else:
    st.warning("📎 Please provide a full path to a PST file.")
