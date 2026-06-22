import os, json, base64, tempfile, mimetypes
from datetime import datetime
from flask import Flask
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
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
    if not token_json:
        raise Exception("GOOGLE_TOKEN_JSON missing in Render Environment")

    creds = Credentials.from_authorized_user_info(json.loads(token_json), SCOPES)
    return build("gmail", "v1", credentials=creds)


def today_text():
    return datetime.now().strftime("%d.%m.%Y")


def today_subject_date():
    return datetime.now().strftime("%d-%b-%Y")


def clean_text_series(s):
    return (
        s.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
        .str.upper()
    )


def download_attachment(service, msg_id, filename):
    msg = service.users().messages().get(
        userId="me",
        id=msg_id
    ).execute()

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


def draft_exists_today(service):
    subject = "Current Jobs Pending List Report as on - " + today_subject_date()

    drafts = service.users().drafts().list(
        userId="me",
        maxResults=50
    ).execute()

    for d in drafts.get("drafts", []):
        draft = service.users().drafts().get(
            userId="me",
            id=d["id"]
        ).execute()

        headers = draft["message"]["payload"].get("headers", [])

        for h in headers:
            if h.get("name", "").lower() == "subject":
                if h.get("value", "").strip() == subject:
                    return True

    return False


def create_draft(service, subject, html_body, attachment_file=None):
    html_body = transform(html_body)

    message = MIMEMultipart()
    message["Subject"] = subject
    message.attach(MIMEText(html_body, "html", "utf-8"))

    if attachment_file and os.path.exists(attachment_file):
        ctype, encoding = mimetypes.guess_type(attachment_file)

        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"

        maintype, subtype = ctype.split("/", 1)

        with open(attachment_file, "rb") as f:
            part = MIMEBase(maintype, subtype)
            part.set_payload(f.read())

        encoders.encode_base64(part)

        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=os.path.basename(attachment_file)
        )

        message.attach(part)

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
    html = """
    <table style="
        border-collapse:collapse;
        font-family:Calibri;
        font-size:13px;
        table-layout:fixed;
    ">
    """

    html += "<thead><tr>"

    for i, col in enumerate(df.columns):
        width = "27ch" if i == 0 else "10.43ch"

        html += f"""
        <th style="
            border:1px solid #A6A6A6;
            background-color:#4F81BD;
            color:white;
            font-weight:bold;
            text-align:center;
            padding:4px;
            width:{width};
            min-width:{width};
            max-width:{width};
        ">{col}</th>
        """

    html += "</tr></thead><tbody>"

    for _, row in df.iterrows():
        first_value = str(row.iloc[0]).strip()
        is_grand_total = first_value.lower() == "grand total"

        html += "<tr>"

        for i, value in enumerate(row):
            col_name = df.columns[i]
            width = "27ch" if i == 0 else "10.43ch"

            style = f"""
                border:1px solid #A6A6A6;
                padding:4px;
                width:{width};
                min-width:{width};
                max-width:{width};
                text-align:center;
                vertical-align:middle;
            """

            if i == 0:
                style += "text-align:left;font-weight:bold;"

            if is_grand_total:
                style += """
                    background-color:#4F81BD;
                    color:white;
                    font-weight:bold;
                """

            if col_name in ["15 to 29", "30 & Above"] and not is_grand_total:
                try:
                    num = int(value)
                except:
                    num = 0

                if num > 0:
                    style += """
                        background-color:red;
                        color:white;
                        font-weight:bold;
                    """

            if pd.isna(value) or value == 0:
                display_value = ""
            else:
                display_value = value

            html += f"<td style='{style}'>{display_value}</td>"

        html += "</tr>"

    html += "</tbody></table>"
    return html


def simple_table_style(df):
    html = """
    <table style="
        border-collapse:collapse;
        font-family:Calibri;
        font-size:13px;
        table-layout:fixed;
    ">
    """

    html += "<thead><tr>"

    for i, col in enumerate(df.columns):
        width = "27ch" if i == 0 else "14ch"

        html += f"""
        <th style="
            border:1px solid #A6A6A6;
            background-color:#4F81BD;
            color:white;
            font-weight:bold;
            text-align:center;
            padding:4px;
            width:{width};
            min-width:{width};
            max-width:{width};
        ">{col}</th>
        """

    html += "</tr></thead><tbody>"

    for _, row in df.iterrows():
        first_value = str(row.iloc[0]).strip()
        is_grand_total = first_value.lower() == "grand total"

        html += "<tr>"

        for i, value in enumerate(row):
            width = "27ch" if i == 0 else "14ch"

            style = f"""
                border:1px solid #A6A6A6;
                padding:4px;
                width:{width};
                min-width:{width};
                max-width:{width};
                text-align:center;
                vertical-align:middle;
            """

            if i == 0:
                style += "text-align:left;font-weight:bold;"

            if is_grand_total:
                style += """
                    background-color:#4F81BD;
                    color:white;
                    font-weight:bold;
                """

            if pd.isna(value) or value == 0:
                display_value = ""
            else:
                display_value = value

            html += f"<td style='{style}'>{display_value}</td>"

        html += "</tr>"

    html += "</tbody></table>"
    return html


def load_clean_files(pending_file, location_file):
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

    return df, loc


def build_pivot_html_from_dataframe(df, loc, intro_text):
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

    df = df.copy()
    loc = loc.copy()

    df["_location_key"] = clean_text_series(df[pending_location_col])
    loc["_branch_key"] = clean_text_series(loc[loc_branch_col])

    state_map = (
        loc.drop_duplicates("_branch_key")
        .set_index("_branch_key")[loc_state_col]
        .to_dict()
    )

    df["Zone"] = (
        df[pending_zone_col]
        .astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
    )

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

    return f"""
    <p style='font-family:Calibri;font-size:14px;'>
    Hi Team,<br><br>
    {intro_text}
    </p>

    <p style='font-family:Calibri;font-size:14px;font-weight:bold;'>Zone Summary:</p>
    {table_style(zone_pivot)}
    <br>

    <p style='font-family:Calibri;font-size:14px;font-weight:bold;'>State Summary:</p>
    {table_style(state_pivot)}
    <br>

    <p style='font-family:Calibri;font-size:14px;'>
    Regards,<br>
    Service Department
    </p>
    """


def build_current_report_html(pending_file, location_file):
    df, loc = load_clean_files(pending_file, location_file)

    return build_pivot_html_from_dataframe(
        df,
        loc,
        "Please find the Zone and State Current Jobs Pending List Report:"
    )


def build_onsite_report_html(pending_file, location_file):
    df, loc = load_clean_files(pending_file, location_file)

    onsite_col = "Onsite Job(Y/N)"

    if onsite_col not in df.columns:
        raise Exception(f"Pending sheet column missing: {onsite_col}")

    df = df[df[onsite_col].astype(str).str.strip().str.upper() == "Y"]

    if df.empty:
        return """
        <p style='font-family:Calibri;font-size:14px;'>
        Hi Team,<br><br>
        No Onsite Pending Jobs found today.<br><br>
        Regards,<br>
        Service Department
        </p>
        """

    return build_pivot_html_from_dataframe(
        df,
        loc,
        "Please find the Zone and State Onsite Jobs Pending Report:"
    )


def build_closure_pending_report(pending_file, location_file):
    df, loc = load_clean_files(pending_file, location_file)

    status_col = "Status"

    if status_col not in df.columns:
        raise Exception(f"Pending sheet column missing: {status_col}")

    df = df[
        df[status_col]
        .astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
        .str.upper()
        == "CLOSURE PENDING"
    ]

    if df.empty:
        html = """
        <p style='font-family:Calibri;font-size:14px;'>
        Hi Team,<br><br>
        No Closure Pending jobs found today.<br><br>
        Regards,<br>
        Service Department
        </p>
        """
        return html, None

    if "Zone" not in df.columns:
        raise Exception("Zone column missing for Closure Pending report")

    if "Location" not in df.columns:
        raise Exception("Location column missing for Closure Pending report")

    if "JOB NO." not in df.columns:
        raise Exception("JOB NO. column missing for Closure Pending report")

    zone_pivot = pd.pivot_table(
        df,
        index="Zone",
        values="JOB NO.",
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    ).reset_index()

    zone_pivot.columns = ["Zone", "Count of JOB NO."]

    location_pivot = pd.pivot_table(
        df,
        index="Location",
        values="JOB NO.",
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    )

    grand = location_pivot.loc[["Grand Total"]]
    normal = location_pivot.drop(index="Grand Total", errors="ignore")
    normal = normal.sort_values(by="JOB NO.", ascending=False)
    location_pivot = pd.concat([normal, grand]).reset_index()
    location_pivot.columns = ["Location", "Count of JOB NO."]

    html = f"""
    <p style='font-family:Calibri;font-size:14px;'>
    Hi Team,<br><br>
    Please find the Closure Pending Report:
    </p>

    <p style='font-family:Calibri;font-size:14px;font-weight:bold;'>Zone Summary:</p>
    {simple_table_style(zone_pivot)}
    <br>

    <p style='font-family:Calibri;font-size:14px;font-weight:bold;'>Location Summary:</p>
    {simple_table_style(location_pivot)}
    <br>

    <p style='font-family:Calibri;font-size:14px;'>
    Regards,<br>
    Service Department
    </p>
    """

    attachment_name = "Closure Pending_" + today_subject_date() + ".xlsx"
    attachment_path = os.path.join(tempfile.gettempdir(), attachment_name)

    with pd.ExcelWriter(attachment_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Closure Pending", index=False)

        zone_pivot.to_excel(
            writer,
            sheet_name="Pivot Closure Pending",
            index=False,
            startrow=2
        )

        location_pivot.to_excel(
            writer,
            sheet_name="Pivot Closure Pending",
            index=False,
            startrow=12
        )

    return html, attachment_path


@app.route("/")
def home():
    return "Pending report automation running"


@app.route("/check-pending-report")
def check_pending_report():
    try:
        service = get_service()
        
        if draft_exists_today(service):
            return "Today's drafts already created. Skipped."

        pending_file, pending_subject, pending_filename = find_pending_file(service)

        if not pending_file:
            return f"Pending file not found: {pending_filename}"

        location_file, location_filename = find_latest_location_file(service)

        if not location_file:
            return f"Location file not found: {location_filename}"

        current_html = build_current_report_html(pending_file, location_file)
        current_subject = "Current Jobs Pending List Report as on - " + today_subject_date()
        current_draft_id = create_draft(service, current_subject, current_html)

        onsite_html = build_onsite_report_html(pending_file, location_file)
        onsite_subject = "Onsite Jobs Pending Report as on - " + today_subject_date()
        onsite_draft_id = create_draft(service, onsite_subject, onsite_html)

        closure_html, closure_attachment = build_closure_pending_report(
            pending_file,
            location_file
        )

        closure_subject = "Closure Pending Report - " + today_subject_date()

        closure_draft_id = create_draft(
            service,
            closure_subject,
            closure_html,
            closure_attachment
        )

        return (
            f"Done. Current Draft: {current_draft_id} | "
            f"Onsite Draft: {onsite_draft_id} | "
            f"Closure Draft: {closure_draft_id}"
        )

    except Exception as e:
        return f"ERROR: {str(e)}", 500
