"""
Event participation form -> saves every submission to a Google Sheet (and
optionally uploads passport/visa copies to a Google Drive folder).

Run locally:  streamlit run app.py
"""
import io
import re
import datetime as dt

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

# ----------------------------------------------------------------------------
# 1. EVENT SETTINGS  (edit these for each event)
# ----------------------------------------------------------------------------
EVENT_NAME = "COP 31"
EVENT_START = dt.date(2026, 11, 9)
EVENT_END = dt.date(2026, 11, 20)
WORKSHEET_NAME = "Registrations"
MAX_UPLOAD_MB = 5

# Declarations taken from your Excel "घोषणा चेकलिस्ट" sheet.
DECLARATIONS = [
    "नेपाल, आयोजक राष्ट्र र UNFCCC सचिवालयका कानुन तथा आचारसंहिता पालना गर्ने",
    "राष्ट्रिय सम्पर्क अधिकृतको पूर्ण सहमतिमा लिखित रूपमा मात्र नेपालको धारणा राख्ने",
    "अन्य राष्ट्र, संस्था वा प्रतिनिधिलाई लक्षित नगरी तथ्यपरक तथा कूटनीतिक धारणा राख्ने",
    "सम्मेलनका सबै फोरममा राखिएको आधिकारिक धारणासँग एकरूपता कायम गर्ने",
    "सम्मेलनको सम्पूर्ण खर्च आफैँ वा उल्लेखित संस्थाले व्यहोर्ने",
    "नेपालले दर्ता गरेको पार्टीको Badge प्रयोग गरी प्रदर्शनमा भाग नलिने",
    "आधिकारिक धारणाविपरीत वा अनावश्यक क्रियाकलाप भए नाम दर्ता खारेज हुन सक्ने कुरा स्वीकार गर्ने",
    "सम्मेलन अवधिभर राष्ट्रिय सम्पर्क अधिकृतसँग समन्वय गर्ने",
    "सम्मेलन सम्पन्न भएको दुई हप्ताभित्र प्रगति प्रतिवेदन बुझाउने",
    "दिइएको विवरण सत्य भएको र झुटा ठहरिए प्रचलित कानुनअनुसार कारबाही स्वीकार गर्ने",
]

# Column order in the Google Sheet (matches your Excel, plus a few extras).
HEADERS = [
    "Timestamp",
    "सम्मेलनको नाम", "सम्मेलन सुरु मिति", "सम्मेलन समाप्त मिति",
    "नाम र थर", "मोबाइल नम्बर", "इमेल", "पद", "संस्था", "ठेगाना",
    "पासपोर्ट नं.", "पासपोर्ट जारी मिति", "पासपोर्ट म्याद सकिने मिति",
    "पासपोर्ट प्रतिलिपि (छ/छैन)", "भिसा प्रतिलिपि (छ/छैन)",
    "हवाई टिकट नं.", "होटलको ठेगाना",
    "खर्च व्यहोर्ने व्यक्ति/संस्था", "खर्च सुनिश्चितता पत्र (छ/छैन)",
    "संस्थाको स्रोत सुनिश्चितता (छ/छैन)",
    "अपलोड गरिएका फाइल (Drive links)",
    "घोषणा स्वीकृत", "नाम दर्ता स्थिति", "कैफियत",
]
EMAIL_COL = HEADERS.index("इमेल") + 1  # 1-based, for duplicate check
EVENT_COL = HEADERS.index("सम्मेलनको नाम") + 1

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ----------------------------------------------------------------------------
# 2. GOOGLE CONNECTIONS
# ----------------------------------------------------------------------------
@st.cache_resource
def _credentials():
    return Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )


@st.cache_resource
def get_worksheet():
    client = gspread.authorize(_credentials())
    book = client.open_by_key(st.secrets["app"]["SHEET_ID"])
    try:
        ws = book.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(WORKSHEET_NAME, rows=500, cols=len(HEADERS))
    if ws.row_values(1) != HEADERS:
        ws.update([HEADERS], "A1")
        ws.freeze(rows=1)
    return ws


def upload_to_drive(file, prefix):
    """Upload one Streamlit UploadedFile to the Drive folder; return its link."""
    drive = build("drive", "v3", credentials=_credentials(), cache_discovery=False)
    meta = {
        "name": f"{prefix}_{file.name}",
        "parents": [st.secrets["app"]["DRIVE_FOLDER_ID"]],
    }
    media = MediaIoBaseUpload(
        io.BytesIO(file.getvalue()),
        mimetype=file.type or "application/octet-stream",
    )
    created = drive.files().create(
        body=meta, media_body=media, fields="webViewLink", supportsAllDrives=True
    ).execute()
    return created["webViewLink"]


# ----------------------------------------------------------------------------
# 3. HELPERS
# ----------------------------------------------------------------------------
def yes_no(label, key, help=None):
    choice = st.radio(label, ["छ (Yes)", "छैन (No)"], horizontal=True,
                      key=key, help=help, index=1)
    return "छ" if choice.startswith("छ ") else "छैन"


def valid_email(value):
    return re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", value or "") is not None


def valid_phone(value):
    digits = re.sub(r"[\s\-+()]", "", value or "")
    return digits.isdigit() and 7 <= len(digits) <= 15


def already_registered(ws, email):
    emails = ws.col_values(EMAIL_COL)[1:]
    events = ws.col_values(EVENT_COL)[1:]
    return any(e.strip().lower() == email.lower() and ev == EVENT_NAME
               for e, ev in zip(emails, events))


# ----------------------------------------------------------------------------
# 4. THE FORM
# ----------------------------------------------------------------------------
st.set_page_config(page_title=f"{EVENT_NAME} Registration", page_icon="📝")
st.title(f"📝 {EVENT_NAME} — सहभागी दर्ता फारम")
st.caption(
    f"{EVENT_NAME} · {EVENT_START:%d %b %Y} – {EVENT_END:%d %b %Y}. "
    "Fields marked * are required / * चिन्ह भएका विवरण अनिवार्य छन्।"
)

uploads_enabled = bool(st.secrets.get("app", {}).get("DRIVE_FOLDER_ID", ""))

with st.form("registration"):
    st.subheader("१. व्यक्तिगत विवरण / Personal details")
    name = st.text_input("नाम र थर / Full name *")
    c1, c2 = st.columns(2)
    mobile = c1.text_input("मोबाइल नम्बर / Mobile *")
    email = c2.text_input("इमेल / Email *")
    c3, c4 = st.columns(2)
    position = c3.text_input("पद / Position *")
    organization = c4.text_input("संस्था / Organization *")
    address = st.text_area("ठेगाना / Address", height=70)

    st.subheader("२. पासपोर्ट र भिसा / Passport & visa")
    passport_no = st.text_input("पासपोर्ट नं. / Passport number")
    d1, d2 = st.columns(2)
    issue = d1.date_input("जारी मिति / Issue date", value=None,
                          min_value=dt.date(2000, 1, 1), max_value=dt.date.today())
    expiry = d2.date_input("म्याद सकिने मिति / Expiry date", value=None,
                           min_value=dt.date(2000, 1, 1), max_value=dt.date(2060, 12, 31))
    e1, e2 = st.columns(2)
    with e1:
        has_passport_copy = yes_no("पासपोर्ट प्रतिलिपि छ?", "pp")
    with e2:
        has_visa_copy = yes_no("भिसा प्रतिलिपि छ?", "visa")

    st.subheader("३. यात्रा र खर्च / Travel & funding")
    flight = st.text_input("हवाई टिकट नं. / Flight ticket no. (airline)")
    hotel = st.text_input("होटलको ठेगाना / Hotel address")
    payer = st.text_input("खर्च व्यहोर्ने व्यक्ति/संस्था / Who covers the cost? *")
    f1, f2 = st.columns(2)
    with f1:
        has_funding_letter = yes_no("खर्च सुनिश्चितता पत्र छ?", "fl")
    with f2:
        has_org_support = yes_no("संस्थाको स्रोत सुनिश्चितता छ?", "org")

    uploaded = {}
    if uploads_enabled:
        st.subheader("४. कागजात अपलोड (वैकल्पिक) / Upload documents (optional)")
        st.caption(f"PDF / JPG / PNG, max {MAX_UPLOAD_MB} MB each")
        types = ["pdf", "jpg", "jpeg", "png"]
        uploaded["Passport"] = st.file_uploader("Passport copy", type=types)
        uploaded["Visa"] = st.file_uploader("Visa copy", type=types)
        uploaded["Funding"] = st.file_uploader("Funding letter", type=types)

    st.subheader("घोषणा / Declaration — सबै स्वीकार गर्नुपर्नेछ")
    agreed = [st.checkbox(f"{i}. {text}", key=f"decl{i}")
              for i, text in enumerate(DECLARATIONS, 1)]

    submitted = st.form_submit_button("दर्ता गर्नुहोस् / Submit", type="primary")

# ----------------------------------------------------------------------------
# 5. VALIDATE + SAVE
# ----------------------------------------------------------------------------
if submitted:
    errors = []
    if not name.strip():
        errors.append("नाम र थर आवश्यक छ।")
    if not valid_phone(mobile):
        errors.append("मान्य मोबाइल नम्बर राख्नुहोस्।")
    if not valid_email(email):
        errors.append("मान्य इमेल राख्नुहोस्।")
    if not position.strip() or not organization.strip():
        errors.append("पद र संस्था आवश्यक छ।")
    if not payer.strip():
        errors.append("खर्च व्यहोर्ने व्यक्ति/संस्था उल्लेख गर्नुहोस्।")
    if issue and expiry and expiry <= issue:
        errors.append("पासपोर्ट म्याद सकिने मिति जारी मितिपछि हुनुपर्छ।")
    if expiry and expiry < EVENT_END + dt.timedelta(days=183):
        st.warning("⚠️ Your passport expires within 6 months of the event end date; "
                   "many countries refuse entry in this case.")
    if not all(agreed):
        errors.append("सबै घोषणा स्वीकार गर्नुपर्नेछ।")
    for label, f in uploaded.items():
        if f and f.size > MAX_UPLOAD_MB * 1024 * 1024:
            errors.append(f"{label}: file is larger than {MAX_UPLOAD_MB} MB.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    try:
        ws = get_worksheet()
        if already_registered(ws, email.strip()):
            st.error("This email is already registered for this event. "
                     "यो इमेलबाट पहिले नै दर्ता भइसकेको छ।")
            st.stop()

        links = []
        safe = re.sub(r"[^A-Za-z0-9]+", "_", name.strip())[:40]
        for label, f in uploaded.items():
            if f:
                links.append(f"{label}: " + upload_to_drive(f, f"{EVENT_NAME}_{safe}_{label}"))

        row = [
            dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            EVENT_NAME, EVENT_START.isoformat(), EVENT_END.isoformat(),
            name.strip(), mobile.strip(), email.strip(),
            position.strip(), organization.strip(), address.strip(),
            passport_no.strip(),
            issue.isoformat() if issue else "",
            expiry.isoformat() if expiry else "",
            has_passport_copy, has_visa_copy,
            flight.strip(), hotel.strip(),
            payer.strip(), has_funding_letter, has_org_support,
            "\n".join(links),
            f"हो ({len(DECLARATIONS)}/{len(DECLARATIONS)})",
            "प्रक्रियामा",   # admin updates this later in the sheet
            "",              # कैफियत (admin remarks)
        ]
        # USER_ENTERED is NOT used, so values like 0098... stay as text
        ws.append_row(row, value_input_option="RAW")
        st.success("✅ तपाईंको दर्ता सफलतापूर्वक पेश भयो। / Your registration was submitted.")
        st.balloons()
    except Exception as exc:  # show a friendly message, keep details in logs
        st.error("Sorry, your registration could not be saved. Please try again "
                 "or contact the organizer.")
        st.exception(exc) if st.secrets.get("app", {}).get("DEBUG") else None
