import smtplib
import ssl
from email.message import EmailMessage
from typing import List, Optional, Dict
import mimetypes
import os


def _guess_mime_type(file_path: str):
    mime_type, _ = mimetypes.guess_type(file_path)
    if not mime_type:
        return ("application", "octet-stream")
    major, minor = mime_type.split("/", 1)
    return (major, minor)


def send_results_email(
    recipients: List[str],
    subject: str,
    body_text: str,
    attachments: Optional[List[str]] = None,
    smtp: Optional[Dict] = None,
    sender: Optional[str] = None,
):
    if not recipients:
        print("✉️ No recipients configured. Skipping email.")
        return

    smtp = smtp or {}
    host = smtp.get("host")
    port = int(smtp.get("port", 587))
    username = smtp.get("username")
    password = smtp.get("password")
    use_tls = smtp.get("use_tls", True)
    use_ssl = smtp.get("use_ssl", False)
    from_addr = sender or smtp.get("from") or username

    if not host or not from_addr:
        print("✉️ SMTP host or sender not configured. Skipping email.")
        return

    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body_text)

    for path in (attachments or []):
        if not path or not os.path.exists(path):
            continue
        maintype, subtype = _guess_mime_type(path)
        with open(path, "rb") as f:
            data = f.read()
        filename = os.path.basename(path)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)

    try:
        if use_ssl and not use_tls:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(host, port, context=context) as server:
                if username and password:
                    server.login(username, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(host, port) as server:
                if use_tls:
                    server.starttls(context=ssl.create_default_context())
                if username and password:
                    server.login(username, password)
                server.send_message(msg)
        print("✉️ Results email sent successfully.")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")






