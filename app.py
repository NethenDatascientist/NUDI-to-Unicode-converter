from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
import functools
import json
import os
import re
import shutil
import smtplib
import subprocess
import tempfile
import threading
import urllib.request

import docx
from kannada_converter import ascii_to_unicode
import openpyxl
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Kannada Unicode Converter", page_icon="🔤", layout="centered"
)

st.title("Kannada ASCII to Unicode Converter")
st.markdown("**GitHub account : Nethen**")
st.write(
    "Upload your `.doc`, `.docx`, `.xls`, or `.xlsx` file to convert it to Unicode format."
)

# --- PRE-COMPILED REGEXES FOR MAXIMUM SPEED ---
EMAIL_RE = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
URL_RE = re.compile(r"https?://\S+|www\.\S+")
SPLIT_RE = re.compile(r"(\s+)")
ACRONYM_OR_NUM_RE = re.compile(r"^[A-Z0-9\-\./]+$|^[0-9\-\./]+$")
ASCII_CHECK_RE = re.compile(r"[\x21-\x7E]")


def get_secret(key, default=None):
    """Reads from Render Environment Variables first, then falls back to st.secrets."""
    val = os.environ.get(key)
    if val:
        return val
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


# --- NON-BLOCKING VISITOR TRACKER & EMAIL ALERT ---
def capture_visitor_context():
    """Captures Streamlit context on the main thread before handing off to background thread."""
    headers = getattr(st.context, "headers", {}) or {}
    forwarded = headers.get("X-Forwarded-For", "")
    ip = (
        forwarded.split(",")[0].strip()
        if forwarded
        else (getattr(st.context, "ip_address", None) or "Unknown")
    )
    user_agent = headers.get("User-Agent", "Unknown Device/Browser")
    platform = headers.get("Sec-Ch-Ua-Platform", "Unknown OS").strip('"')
    tz = getattr(st.context, "timezone", None) or "Unknown Timezone"
    locale = getattr(st.context, "locale", None) or headers.get(
        "Accept-Language", "Unknown"
    )
    return ip, user_agent, platform, tz, locale


def _send_alert_background(
    event_label,
    ip,
    user_agent,
    platform,
    tz,
    locale,
    sender_email,
    app_password,
    receiver_email,
):
    """Runs IP lookup and SMTP email in the background so the UI never freezes."""
    try:
        location_str = "Location unavailable"
        isp_str = "Unknown ISP"
        if ip and ip not in ["Unknown", "127.0.0.1", "::1"]:
            try:
                url = f"http://ip-api.com/json/{ip}?fields=status,country,regionName,city,isp"
                req = urllib.request.Request(
                    url, headers={"User-Agent": "Mozilla/5.0"}
                )
                with urllib.request.urlopen(req, timeout=2) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    if data.get("status") == "success":
                        location_str = f"{data.get('city', '')}, {data.get('regionName', '')}, {data.get('country', '')}"
                        isp_str = data.get("isp", "Unknown ISP")
            except Exception:
                pass

        ist_now = datetime.now(timezone(timedelta(hours=5, minutes=30))).strftime(
            "%Y-%m-%d %I:%M:%S %p IST"
        )
        info = (
            f"• Time: {ist_now}\n"
            f"• IP Address: {ip}\n"
            f"• Location: {location_str}\n"
            f"• ISP / Network: {isp_str}\n"
            f"• Platform: {platform}\n"
            f"• Timezone: {tz}\n"
            f"• Language: {locale}\n"
            f"• User Agent: {user_agent}\n"
        )

        print(
            f"\n--- 🔔 VISITOR TRACKER ({event_label}) ---\n{info}------------------------------------------",
            flush=True,
        )

        if sender_email and app_password:
            msg = EmailMessage()
            msg["Subject"] = f"🔔 Kannada Converter: {event_label}"
            msg["From"] = sender_email
            msg["To"] = receiver_email
            msg.set_content(
                f"Hello Nethen,\n\nActivity detected on your Kannada ASCII to Unicode"
                f" Converter ({event_label}):\n\n{info}"
            )
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=5) as server:
                server.login(sender_email, app_password)
                server.send_message(msg)
    except Exception:
        pass


def notify_visit(event_label="Opened the App"):
    try:
        ip, user_agent, platform, tz, locale = capture_visitor_context()
        sender_email = get_secret("NOTIFY_EMAIL")
        app_password = get_secret("NOTIFY_APP_PASSWORD")
        receiver_email = get_secret("RECEIVER_EMAIL", "dsdanethen@gmail.com")

        threading.Thread(
            target=_send_alert_background,
            args=(
                event_label,
                ip,
                user_agent,
                platform,
                tz,
                locale,
                sender_email,
                app_password,
                receiver_email,
            ),
            daemon=True,
        ).start()
    except Exception:
        pass


if "visit_tracked" not in st.session_state:
    notify_visit("Opened the App Link")
    st.session_state.visit_tracked = True


# --- COMMON ENGLISH WORDS & ADMINISTRATIVE TERMS TO PRESERVE ---
COMMON_ENGLISH_WORDS = {
    "for", "and", "the", "of", "to", "in", "on", "with", "by", "from", "as", "at", "an",
    "is", "are", "was", "were", "be", "been", "being", "or", "not", "but", "if", "than",
    "then", "so", "it", "its", "this", "that", "these", "those", "which", "who", "whom",
    "whose", "where", "when", "why", "how", "all", "each", "every", "both", "few",
    "more", "most", "other", "some", "such", "no", "nor", "too", "very", "can", "will",
    "just", "should", "now", "also", "into", "over", "after", "up", "out", "down", "off",
    "again", "further", "once", "here", "there", "any", "only", "own", "same", "central",
    "state", "share", "scheme", "project", "total", "grant", "amount", "district",
    "taluk", "irrigation", "department", "works", "minor", "major", "estimated", "cost",
    "financial", "year", "years", "lakhs", "crores", "rupees", "rs", "nos", "no", "sl",
    "date", "page", "memo", "reference", "subject", "description", "sub", "para", "note",
    "report", "file", "document", "pmksy", "smi", "water", "resource", "resources",
    "canal", "canals", "tank", "tanks", "lake", "lakes", "bridge", "bridges", "road",
    "roads", "building", "buildings", "office", "officer", "engineer", "executive",
    "assistant", "junior", "chief", "superintending", "secretary", "minister", "mla",
    "mlc", "honble", "assembly", "council", "question", "answer", "annexure",
    "karnataka", "kalaburagi", "yadgir", "gulbarga",
}

MANUAL_FIXES = {"¯ï": "ಲ್", "¯i": "ಲ್"}


# --- TWO-LEVEL CACHED TRANSLITERATION (WORD + SENTENCE) ---
@functools.lru_cache(maxsize=50000)
def convert_single_token(token):
    """Caches individual words so repeated terms across thousands of cells convert instantly."""
    stripped = token.strip()
    if not stripped or not ASCII_CHECK_RE.search(stripped):
        return token

    cleaned_token = stripped.strip(".,;:!?()[]{}'\"-").lower()
    if (
        ACRONYM_OR_NUM_RE.match(stripped)
        or cleaned_token in COMMON_ENGLISH_WORDS
    ):
        return token

    try:
        conv = ascii_to_unicode(token)
        if conv:
            for ascii_bug, unicode_fix in MANUAL_FIXES.items():
                if ascii_bug in conv:
                    conv = conv.replace(ascii_bug, unicode_fix)
            return conv
        return token
    except Exception:
        return token


@functools.lru_cache(maxsize=20000)
def transliterate_text_safely(text):
    if not isinstance(text, str) or not text.strip():
        return text

    # Skip strings that contain no ASCII characters (e.g., already pure Kannada Unicode)
    if not ASCII_CHECK_RE.search(text):
        return text

    emails = EMAIL_RE.findall(text)
    urls = URL_RE.findall(text)
    protected_items = emails + urls
    temp_text = text

    for i, item in enumerate(protected_items):
        temp_text = temp_text.replace(item, f"_PROTECTED{i}__")

    tokens = SPLIT_RE.split(temp_text)
    translated_tokens = [
        token
        if (token.isspace() or token.startswith("_PROTECTED"))
        else convert_single_token(token)
        for token in tokens
    ]

    converted = "".join(translated_tokens)

    for i, item in enumerate(protected_items):
        converted = converted.replace(f"_PROTECTED{i}__", item)

    return converted


# --- FAST & ISOLATED LIBREOFFICE CONVERTER ---
def convert_with_libreoffice(input_file, target_ext):
    out_dir = os.path.dirname(os.path.abspath(input_file))
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        env = os.environ.copy()
        env["HOME"] = "/tmp"

        subprocess.run(
            [
                "soffice",
                "--headless",
                "--invisible",
                "--nodefault",
                "--nologo",
                "--nolockcheck",
                "--nofirststartwizard",
                "--norestore",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to",
                target_ext,
                "--outdir",
                out_dir,
                input_file,
            ],
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=120,
        )
    except subprocess.TimeoutExpired:
        st.error(
            "⚠️ File conversion timed out. Please save the file as `.xlsx` or `.docx` on your computer and upload again."
        )
        return None
    except Exception as e:
        st.error(f"⚠️ Conversion error: {e}")
        return None
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    output_file = f"{os.path.splitext(input_file)[0]}.{target_ext}"
    return output_file if os.path.exists(output_file) else None


def process_word_document(input_path, output_path):
    doc = docx.Document(input_path)

    for para in doc.paragraphs:
        if para.text and para.text.strip():
            translated = transliterate_text_safely(para.text)
            para.clear()
            new_run = para.add_run(translated)
            new_run.font.name = "Tunga"

    # Track XML cell elements so merged cells are only processed once
    seen_cells = set()
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell._tc in seen_cells:
                    continue
                seen_cells.add(cell._tc)
                for para in cell.paragraphs:
                    if para.text and para.text.strip():
                        translated = transliterate_text_safely(para.text)
                        para.clear()
                        new_run = para.add_run(translated)
                        new_run.font.name = "Tunga"

    doc.save(output_path)


def process_excel_document(input_path, output_path):
    wb = openpyxl.load_workbook(input_path)

    for sheet in wb.worksheets:
        # Iterate ONLY over populated cells stored in memory, skipping millions of empty grid cells
        cells_dict = getattr(sheet, "_cells", None)
        cell_iterable = (
            cells_dict.values()
            if cells_dict is not None
            else (c for row in sheet.iter_rows() for c in row)
        )

        for cell in cell_iterable:
            val = cell.value
            if isinstance(val, str):
                val_str = val.strip()
                # Never touch Excel formulas
                if val_str and not val_str.startswith("="):
                    cell.value = transliterate_text_safely(val)

    wb.save(output_path)


# --- FILE UPLOAD & PROCESSING ---
uploaded_file = st.file_uploader(
    "ಫೈಲ್ ಅನ್ನು ಆಯ್ಕೆ ಮಾಡಿ (Choose file)", type=["doc", "docx", "xls", "xlsx"]
)

if uploaded_file is not None:
    filename = uploaded_file.name

    if st.session_state.get("last_tracked_file") != filename:
        notify_visit(f"Uploaded File: {filename}")
        st.session_state.last_tracked_file = filename

    with st.spinner("⚙️ Processing and translating text... Please wait."):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, filename)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            base_name, file_ext = os.path.splitext(filename)
            file_ext = file_ext.lower()
            output_ext = ".docx" if file_ext in [".doc", ".docx"] else ".xlsx"
            download_filename = f"kannada_unicode_{base_name}{output_ext}"
            final_output_path = os.path.join(tmpdir, download_filename)

            success = False

            if file_ext == ".doc":
                converted_path = convert_with_libreoffice(input_path, "docx")
                if converted_path:
                    process_word_document(converted_path, final_output_path)
                    success = True
            elif file_ext == ".docx":
                process_word_document(input_path, final_output_path)
                success = True
            elif file_ext == ".xls":
                converted_path = convert_with_libreoffice(input_path, "xlsx")
                if converted_path:
                    process_excel_document(converted_path, final_output_path)
                    success = True
            elif file_ext == ".xlsx":
                process_excel_document(input_path, final_output_path)
                success = True

            if success and os.path.exists(final_output_path):
                with open(final_output_path, "rb") as f:
                    output_bytes = f.read()

                st.success("✅ Processing complete!")
                st.download_button(
                    label="📥 Download Converted File",
                    data=output_bytes,
                    file_name=download_filename,
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        if output_ext == ".docx"
                        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )
