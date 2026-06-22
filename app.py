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

FROM_MAIL = "sc6@zebcare.in"

PENDING_SHEET = "CustomerPending"
LOCATION_SHEET = "Sheet1"
LOCATION_FILE = "location123456789.xlsx"


def get_service():
    token_json = os.environ.get("GOOGLE_TOKEN_JSON")
    creds = Credentials.from_authorized_user_info(json.loads(token_json), SCOPES)
    return build("gmail", "v1", credentials=creds)


def today_text():
    return datetime.now().strftime("%d.%m.%Y")


def clean_text_series(s):
    return (
        s.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
        .str.upper()
    )


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
        maxResults=10
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

    for msg in result.get("messages", []):
        file_path = download_attachment(service, msg["id"], LOCATION_FILE)
        if file_path:
            return file_path, LOCATION_FILE

    return None, LOCATION_FILE


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


def tat_bucket(days):
    try:
        days = float(days)
    except:
        return "NA"

    if 0 <= days <= 3:
        return "0 to 3"
    elif 4 <= days <= 6:
        return "4 to 6"
    elif 7 <= days <= 10:
        return "7 to 10"
    elif 11 <= days <= 14:
        return "11 to 14"
    elif 15 <= days <= 29:
        return "15 to 29"
    else:
        return "30 & Above"


def table_style(df):
    def highlight(v):
        try:
            v = int(v)
        except:
            return ""

        if v > 0:
            return "background-color:red;color:white;font-weight:bold;text-align:center;"
        return "text-align:center;"

    styled = df.style.set_table_attributes(
        'border="1" cellspacing="0" cellpadding="4"'
    ).set_properties(**{
        "font-family": "Calibri",
        "font-size": "13px",
        "border": "1px solid black",
        "text-align": "center"
    })

    for col in ["30 & Above", "15 to 29"]:
        if col in df.columns:
            styled = styled.map(highlight, subset=[col])

    return styled.hide(axis="index").to_html()


def build_report_html(pending_file, location_file):
    df = pd.read_excel(pending_file, sheet_name=PENDING_SHEET)
    loc = pd.read_excel(location_file, sheet_name=LOCATION_SHEET)

    df.columns = df.columns.astype(str).str.strip()
    loc.columns = loc.columns.astype(str).str.strip()

    df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
    loc = loc.apply(lambda col: col.str.strip() if col.dtype == "object" else col)

    df = df[
        ~df.astype(str)
        .apply(lambda row: row.str.contains("Total Jobs", case=False, na=False))
        .any(axis=1)
    ]

    df = df.dropna(how="all")

    pending_location_col = "Location"
    pending_days_col = "Pending From No. Of Days"
    pending_zone_col = "Zone"

    loc_branch_col = "BRANCH"
    loc_state_col = "State"

    for col in [pending_location_col, pending_days_col, pending_zone_col, "JOB NO."]:
        if col not in df.columns:
            raise Exception(f"Pending sheet column missing: {col}")

    for col in [loc_branch_col, loc_state_col]:
        if col not in loc.columns:
            raise Exception(f"Location sheet column missing: {col}")

    df["_location_key"] = clean_text_series(df[pending_location_col])
    loc["_branch_key"] = clean_text_series(loc[loc_branch_col])

    state_map = (
        loc.drop_duplicates("_branch_key")
        .set_index("_branch_key")[loc_state_col]
        .to_dict()
    )

    df["Zone"] = df[pending_zone_col].astype(str).str.replace("\u00a0", " ", regex=False).str.strip()
    df["State"] = df["_location_key"].map(state_map).fillna("NA")

    df["TAT"] = pd.to_numeric(df[pending_days_col], errors="coerce").fillna(0)
    df["TATX"] = df["TAT"].apply(tat_bucket)

    order_cols = ["0 to 3", "4 to 6", "7 to 10", "11 to 14", "15 to 29", "30 & Above"]

    zone_pivot = pd.pivot_table(
        df,
        index="Zone",
        columns="TATX",
        values="JOB NO.",
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    )

    state_pivot = pd.pivot_table(
        df,
        index="State",
        columns="TATX",
        values="JOB NO.",
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

    zone_pivot.columns.name = None
    state_pivot.columns.name = None

    zone_pivot.index.name = "Zone"
    state_pivot.index.name = "State"

    zone_pivot = zone_pivot.reset_index()
    state_pivot = state_pivot.reset_index()

    zone_html = table_style(zone_pivot)
    state_html = table_style(state_pivot)

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
    try:
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

        return f"Done. Draft created: {draft_id}"

    except Exception as e:
        return f"ERROR: {str(e)}", 500
