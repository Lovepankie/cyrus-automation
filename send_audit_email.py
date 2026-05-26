"""
Send the latest audit merged file as an email attachment via Gmail.
Reads GMAIL_APP_PASSWORD from environment (GitHub Secret).
"""

import os
import smtplib
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

SENDER    = "dkaweesi1@gmail.com"
RECIPIENT = "haddybatte@gmail.com"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

HERE = Path(__file__).parent

def find_latest_merged_file() -> Path:
    marker = HERE / ".last_merged_file"
    if marker.exists():
        path = Path(marker.read_text().strip())
        if path.exists():
            return path

    output_dir = HERE / "merged_output_audit"
    xlsx_files = sorted(output_dir.rglob("audit_merged_*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not xlsx_files:
        raise FileNotFoundError("No merged audit file found.")
    return xlsx_files[-1]


def send_email(file_path: Path, app_password: str):
    today     = datetime.now().strftime("%d %b %Y")
    subject   = f"Daily Audit Report — {today}"
    body      = (
        f"Hi,\n\n"
        f"Please find attached the daily audit report generated on {today}.\n\n"
        f"File: {file_path.name}\n\n"
        f"This is an automated email sent daily at 6:00 AM EAT.\n"
    )

    msg = MIMEMultipart()
    msg["From"]    = SENDER
    msg["To"]      = RECIPIENT
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with open(file_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=file_path.name)
    part["Content-Disposition"] = f'attachment; filename="{file_path.name}"'
    msg.attach(part)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SENDER, app_password)
        server.sendmail(SENDER, RECIPIENT, msg.as_string())

    print(f"Email sent to {RECIPIENT} with attachment: {file_path.name}")


if __name__ == "__main__":
    app_password = os.environ.get("GMAIL_APP_PASSWORD", "").strip()
    if not app_password:
        print("ERROR: GMAIL_APP_PASSWORD environment variable is not set.")
        sys.exit(1)

    file_path = find_latest_merged_file()
    print(f"Attaching: {file_path}")
    send_email(file_path, app_password)
