import os, json, base64, tempfile
from datetime import datetime
from flask import Flask
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from premailer import transform
import pandas as pd

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

app = Flask(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.readonly"
]

FROM_MAIL = "service.support@zebcare.in"

def get_service():
    token_json = os.environ.get("GOOGLE_TOKEN_JSON")
    creds = Credentials.from_authorized_user_info(json.loads(token_json), SCOPES)
    return build("gmail", "v1", credentials=creds)

def today_text():
    return datetime.now().strftime("%d.%m.%Y")

def download_attachment(service, msg_id, filename):
    msg = service.users().messages().get(userId="me", id=msg_id).execute()

    def scan_parts(parts):
        for part in parts:
            if part.get("filename") == filename:
                att_id = part["body"]["attachmentId"]
                att = service.users().messages().attachments().get(
                    userId="me",
                    messageId=msg_id,
                    id=att_id
                ).execute()

                data = base64.urlsafe_b64decode(att["data"])
                path = os.path.join(tempfile.gettempdir(), filename)

                with open(path, "wb") as f:
                    f.write(data)

                return path

            if "parts" in part:
                found = scan_parts(part["parts"])
                if found:
                    return found

        return None

    return scan_parts(msg.get("payload", {}).get("parts", []))

def find_pending_file(service):
    date_text = today_text()
    subject = f"Pending report {date_text}"
    filename = f"Pending report {date_text}.xlsx"

    query = f'from:{FROM_MAIL} subject:"{subject}" has:attachment'

    result = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=5
    ).execute()

    for msg in result.get("messages", []):
        file_path = download_attachment(service, msg["id"], filename)
        if file_path:
            return file_path, subject, filename

    return None, subject, filename

def find_latest_location_file(service):
    query = f'from:{FROM_MAIL} filename:xlsx location123456789'

    result = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=10
    ).execute()

    filename = "location123456789.xlsx"

    for msg in result.get("messages", []):
        file_path = download_attachment(service, msg["id"], filename)
        if file_path:
            return file_path, filename

    return None, filename

def create_draft(service, subject, html_body):
    html_body = transform(html_body)

    message = MIMEMultipart()
    message["Subject"] = subject
    message.attach(MIMEText(html_body, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

    draft = service.users().drafts().create(
        userId="me",
        body={"message": {"raw": raw}}
    ).execute()

    return draft["id"]

def tat_bucket(x):
    try:
        x = float(x)
    except:
        return "NA"

    if 0 <= x <= 3:
        return "0 to 3"
    elif 4 <= x <= 6:
        return "4 to 6"
    elif 7 <= x <= 10:
        return "7 to 10"
    elif 11 <= x <= 14:
        return "11 to 14"
    elif 15 <= x <= 29:
        return "15 to 29"
    else:
        return "30 & Above"

def build_report_html(pending_file, location_file):
    df = pd.read_excel(pending_file, sheet_name="CustomerPending")
    loc = pd.read_excel(location_file, sheet_name="Sheet1")

    # Same as VBA VLOOKUP:
    # Data D column lookup with Sheet1 B:F, return 5th column
    # Here expected:
    # df 4th column = lookup key
    # loc 2nd column = lookup key
    # loc 6th column = Zone
    df_key_col = df.columns[3]
    loc_key_col = loc.columns[1]
    loc_zone_col = loc.columns[5]

    loc_map = loc.set_index(loc_key_col)[loc_zone_col].to_dict()
    df["Zone"] = df[df_key_col].map(loc_map).fillna("NA")

    if "TAT" not in df.columns:
        raise Exception("TAT column not found in CustomerPending sheet")

    if "State" not in df.columns:
        raise Exception("State column not found in CustomerPending sheet")

    df["TATX"] = df["TAT"].apply(tat_bucket)

    order_cols = ["30 & Above", "15 to 29", "11 to 14", "7 to 10", "4 to 6", "0 to 3"]

    zone_pivot = pd.pivot_table(
        df,
        index="Zone",
        columns="TATX",
        values="TAT",
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    )

    state_pivot = pd.pivot_table(
        df,
        index="State",
        columns="TATX",
        values="TAT",
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    )

    zone_cols = [c for c in order_cols if c in zone_pivot.columns]
    state_cols = [c for c in order_cols if c in state_pivot.columns]

    zone_pivot = zone_pivot[zone_cols + ["Grand Total"]]
    state_pivot = state_pivot[state_cols + ["Grand Total"]]

    if "15 to 29" in state_pivot.columns:
        grand = state_pivot.loc[["Grand Total"]]
        normal = state_pivot.drop(index="Grand Total", errors="ignore")
        normal = normal.sort_values(by="15 to 29", ascending=False)
        state_pivot = pd.concat([normal, grand])

    zone_html = zone_pivot.to_html(border=1)
    state_html = state_pivot.to_html(border=1)

    return f"""
    <p style='font-family:Calibri;font-size:14px;'>
    Hi Team,<br><br>
    Please find the Zone and State Current Jobs Pending List Report:
    </p>

    <p style='font-family:Calibri;font-size:14px;font-weight:bold;'>Zone Summary:</p>
    {zone_html}
    <br>

    <p style='font-family:Calibri;font-size:14px;font-weight:bold;'>State Summary:</p>
    {state_html}
    <br>

    <p style='font-family:Calibri;font-size:14px;'>
    Regards,<br>
    Service Department
    </p>
    """

@app.route("/")
def home():
    return "Pending report automation running"

@app.route("/check-pending-report")
def check_pending_report():
    service = get_service()

    pending_file, pending_subject, pending_filename = find_pending_file(service)

    if not pending_file:
        return f"Pending file not found: {pending_filename}"

    location_file, location_filename = find_latest_location_file(service)

    if not location_file:
        return f"Location file not found: {location_filename}"

    html_body = build_report_html(pending_file, location_file)

    draft_subject = "Current Jobs Pending List Report as on - " + datetime.now().strftime("%d-%b-%Y")

    draft_id = create_draft(service, draft_subject, html_body)

    return f"Done. Pending + Location processed. Draft created: {draft_id}"
