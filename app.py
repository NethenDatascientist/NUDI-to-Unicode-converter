import json
import os
import re
import shutil
import smtplib
import subprocess
import threading
import urllib.request
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage

import docx
import openpyxl
import streamlit as st
from kannada_converter import ascii_to_unicode

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Kannada Unicode Converter",
    page_icon="🔤",
    layout="centered",
)

st.title("Kannada ASCII to Unicode Converter")
st.markdown("**GitHub account : Nethen**")
st.write(
    "Upload your `.doc`, `.docx`, or `.xlsx` file to convert it to Unicode format."
)


# -----------------------------------------------------------------------------
# SAFE SECRET / ENVIRONMENT VARIABLE READER FOR RENDER.COM
# -----------------------------------------------------------------------------
def get_secret(key: str, default: str = "") -> str:
    """Reads from Render Environment Variables first, falling back to st.secrets safely."""
    env_val = os.environ.get(key)
    if env_val is not None and str(env_val).strip() != "":
        return str(env_val).strip()
    try:
        sec_val = st.secrets.get(key)
        if sec_val is not None and str(sec_val).strip() != "":
            return str(sec_val).strip()
    except Exception:
        pass
    return default


# -----------------------------------------------------------------------------
# SILENT VISITOR TRACKER & NON-BLOCKING EMAIL ALERT
# -----------------------------------------------------------------------------
def get_visitor_info() -> str:
    """Captures visitor IP, approximate location/ISP, device, and timestamp silently."""
    headers = getattr(st.context, "headers", {}) or {}

    # 1. Get real client IP (prioritize X-Forwarded-For on Render / Cloud proxies)
    forwarded = headers.get("X-Forwarded-For", "") or headers.get("x-forwarded-for", "")
    if forwarded:
        ip = forwarded.split(",")[0].strip()
    else:
        ip = getattr(st.context, "ip_address", None) or "Unknown"

    # 2. Browser & System details
    user_agent = (
        headers.get("User-Agent")
        or headers.get("user-agent")
        or "Unknown Device/Browser"
    )
    platform = (
        headers.get("Sec-Ch-Ua-Platform")
        or headers.get("sec-ch-ua-platform")
        or "Unknown OS"
    ).strip('"')
    tz = getattr(st.context, "timezone", None) or "Unknown Timezone"
    locale = (
        getattr(st.context, "locale", None)
        or headers.get("Accept-Language")
        or headers.get("accept-language")
        or "Unknown"
    )

    # 3. Geolocation lookup from IP
    location_str = "Location unavailable"
    isp_str = "Unknown ISP"
    if ip and ip not in ["Unknown", "127.0.0.1", "::1"]:
        try:
            url = f"http://ip-api.com/json/{ip}?fields=status,country,regionName,city,isp"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
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

    return (
        f"• Time: {ist_now}\n"
        f"• IP Address: {ip}\n"
        f"• Location: {location_str}\n"
        f"• ISP / Network: {isp_str}\n"
        f"• Platform: {platform}\n"
        f"• Timezone: {tz}\n"
        f"• Language: {locale}\n"
        f"• User Agent: {user_agent}\n"
    )


def _send_email_background(
    event_label: str,
    info: str,
    sender_email: str,
    app_password: str,
    receiver_email: str,
):
    """Sends the email in a background thread with a strict timeout so UI never hangs."""
    try:
        msg = EmailMessage()
        msg["Subject"] = f"🔔 Kannada Converter: {event_label}"
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg.set_content(
            f"Hello Nethen,\n\nActivity detected on your Kannada ASCII to Unicode Converter ({event_label}):\n\n{info}"
        )

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=5) as server:
            server.login(sender_email, app_password)
            server.send_message(msg)
    except Exception:
        pass


def notify_visit(event_label: str = "Opened the App"):
    """Prints visitor details to Render logs and triggers a non-blocking email alert."""
    try:
        info = get_visitor_info()
        print(
            f"\n--- 🔔 VISITOR TRACKER ({event_label}) ---\n{info}------------------------------------------",
            flush=True,
        )

        sender_email = get_secret("NOTIFY_EMAIL")
        app_password = get_secret("NOTIFY_APP_PASSWORD")
        receiver_email = get_secret("RECEIVER_EMAIL", "dsdanethen@gmail.com")

        if sender_email and app_password:
            threading.Thread(
                target=_send_email_background,
                args=(event_label, info, sender_email, app_password, receiver_email),
                daemon=True,
            ).start()
    except Exception:
        pass


# Trigger once per session as soon as someone opens the link
if "visit_tracked" not in st.session_state:
    notify_visit("Opened the App Link")
    st.session_state.visit_tracked = True


# -----------------------------------------------------------------------------
# CONVERSION LOGIC
# -----------------------------------------------------------------------------
def transliterate_text_safely(text):
    if not isinstance(text, str) or not text.strip():
        return text

    email_pattern = r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+"
    url_pattern = r"https?://\S+|www\.\S+"

    emails = re.findall(email_pattern, text)
    urls = re.findall(url_pattern, text)

    protected_items = emails + urls
    temp_text = text

    for i, item in enumerate(protected_items):
        temp_text = temp_text.replace(item, f"_PROTECTED{i}__")

    tokens = re.split(r"(\s+)", temp_text)

    translated_tokens = []
    for token in tokens:
        if token.isspace() or token.startswith("_PROTECTED"):
            translated_tokens.append(token)
        else:
            try:
                conv = ascii_to_unicode(token)
                translated_tokens.append(conv if conv else token)
            except Exception:
                translated_tokens.append(token)

    converted = "".join(translated_tokens)

    for i, item in enumerate(protected_items):
        converted = converted.replace(f"_PROTECTED{i}__", item)

    manual_fixes = {"¯ï": "ಲ್", "¯i": "ಲ್"}

    for ascii_bug, unicode_fix in manual_fixes.items():
        converted = converted.replace(ascii_bug, unicode_fix)

    return converted


def convert_doc_to_docx(input_file: str):
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        return None
    try:
        subprocess.run(
            [soffice_bin, "--headless", "--convert-to", "docx", input_file],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=60,
            check=False,
        )
        output_file = f"{os.path.splitext(input_file)[0]}.docx"
        return output_file if os.path.exists(output_file) else None
    except Exception:
        return None


def process_word_document(input_path: str, output_path: str):
    doc = docx.Document(input_path)

    for para in doc.paragraphs:
        if para.text.strip():
            translated = transliterate_text_safely(para.text)
            para.clear()
            new_run = para.add_run(translated)
            new_run.font.name = "Tunga"

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip():
                        translated = transliterate_text_safely(para.text)
                        para.clear()
                        new_run = para.add_run(translated)
                        new_run.font.name = "Tunga"

    doc.save(output_path)


def process_excel_document(input_path: str, output_path: str):
    wb = openpyxl.load_workbook(input_path)

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    cell.value = transliterate_text_safely(cell.value)

    wb.save(output_path)


# -----------------------------------------------------------------------------
# STREAMLIT FILE UPLOADER & DOWNLOAD UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "ಫೈಲ್ ಅನ್ನು ಆಯ್ಕೆ ಮಾಡಿ (Choose file)", type=["doc", "docx", "xlsx"]
)

if uploaded_file is not None:
    filename = os.path.basename(uploaded_file.name)

    # Log and email when a file is uploaded
    if st.session_state.get("last_tracked_file") != filename:
        notify_visit(f"Uploaded File: {filename}")
        st.session_state.last_tracked_file = filename

    # Save uploaded file locally in the container
    with open(filename, "wb") as f:
        f.write(uploaded_file.getbuffer())

    base_name, file_ext = os.path.splitext(filename)
    file_ext = file_ext.lower()
    final_output_name = f"kannada_unicode_{base_name}"
    final_output = None
    temp_docx = None

    with st.spinner("⚙️ Processing and translating text... Please wait."):
        if file_ext == ".doc":
            temp_docx = convert_doc_to_docx(filename)
            if temp_docx:
                final_output = f"{final_output_name}.docx"
                process_word_document(temp_docx, final_output)
            else:
                st.error(
                    "⚠️ Legacy `.doc` conversion requires LibreOffice (`soffice`). "
                    "Please save your file as `.docx` in Word and upload again, or deploy using a Dockerfile with LibreOffice installed."
                )
        elif file_ext == ".docx":
            final_output = f"{final_output_name}.docx"
            process_word_document(filename, final_output)
        elif file_ext == ".xlsx":
            final_output = f"{final_output_name}.xlsx"
            process_excel_document(filename, final_output)

    if final_output and os.path.exists(final_output):
        with open(final_output, "rb") as file_obj:
            converted_bytes = file_obj.read()

        st.success("✅ Processing complete!")
        st.download_button(
            label="📥 Download Converted File",
            data=converted_bytes,
            file_name=final_output,
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                if file_ext in [".doc", ".docx"]
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    # Cleanup all temporary files immediately to keep Render storage clean
    for path_to_remove in [filename, temp_docx, final_output]:
        if path_to_remove and os.path.exists(path_to_remove):
            try:
                os.remove(path_to_remove)
            except Exception:
                pass
