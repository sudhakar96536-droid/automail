import os
import json
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from premailer import transform
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.readonly"
]

FROM_MAIL = "sc6@zebcare.in"
CHECK_SUBJECT = "SAMPLE"

def get_service():
    token_json = os.environ.get("GOOGLE_TOKEN_JSON")
    if not token_json:
        raise Exception("GOOGLE_TOKEN_JSON missing")

    creds = Credentials.from_authorized_user_info(json.loads(token_json), SCOPES)
    return build("gmail", "v1", credentials=creds)

def create_draft(service):
    html_body = """
    <html>
    <body>
        <p>Dear Team,</p>
        <p>This is sample auto draft created from Render Cron Job.</p>
        <p>Mail received from <b>sc6@zebcare.in</b> with subject <b>SAMPLE</b>.</p>
        <p>Regards,<br>Auto Bot</p>
    </body>
    </html>
    """

    html_body = transform(html_body)

    message = MIMEMultipart()
    message["Subject"] = "SAMPLE AUTO DRAFT TEST"
    message.attach(MIMEText(html_body, "html", "utf-8"))

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    body = {
        "message": {
            "raw": raw_message
        }
    }

    draft = service.users().drafts().create(userId="me", body=body).execute()
    print("Draft created:", draft["id"])

def main():
    service = get_service()

    query = f'from:{FROM_MAIL} subject:"{CHECK_SUBJECT}" newer_than:1d'

    result = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=1
    ).execute()

    messages = result.get("messages", [])

    if not messages:
        print("No SAMPLE mail found")
        return

    create_draft(service)

if __name__ == "__main__":
    main()
