import os, json, base64
from flask import Flask
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from premailer import transform
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

app = Flask(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.readonly"
]

FROM_MAIL = "sc6@zebcare.in"
CHECK_SUBJECT = "SAMPLE"

def get_service():
    token_json = os.environ.get("GOOGLE_TOKEN_JSON")
    creds = Credentials.from_authorized_user_info(json.loads(token_json), SCOPES)
    return build("gmail", "v1", credentials=creds)

def create_draft(service):
    html_body = """
    <p>Dear Team,</p>
    <p>This is sample auto draft created from Render.</p>
    <p>Regards,<br>Auto Bot</p>
    """
    html_body = transform(html_body)

    message = MIMEMultipart()
    message["Subject"] = "SAMPLE AUTO DRAFT TEST"
    message.attach(MIMEText(html_body, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

    draft = service.users().drafts().create(
        userId="me",
        body={"message": {"raw": raw}}
    ).execute()

    return draft["id"]

@app.route("/")
def home():
    return "Gmail automation running"

@app.route("/check-sample")
def check_sample():
    service = get_service()

    query = f'from:{FROM_MAIL} subject:"{CHECK_SUBJECT}" newer_than:1d'

    result = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=1
    ).execute()

    if not result.get("messages"):
        return "No SAMPLE mail found"

    draft_id = create_draft(service)
    return f"SAMPLE mail found. Draft created: {draft_id}"
