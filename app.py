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
import openpyxl
import streamlit as st

# Page configuration
# Page configuration
st.set_page_config(
    page_title="Kannada Unicode Converter", page_icon="🔤", layout="centered"
)

# Animated, Eye-Friendly Professional SaaS Theme
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

/* Global Typography */
html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* 1. Remove Top Header & Footer for Zero-Scroll Layout */
[data-testid="stHeader"], footer {
    display: none !important;
}

/* 2. Keyframe Animations */
@keyframes ambientBg {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

@keyframes cardReveal {
    0% { opacity: 0; transform: translateY(18px) scale(0.985); }
    100% { opacity: 1; transform: translateY(0) scale(1); }
}

@keyframes staggerFade {
    0% { opacity: 0; transform: translateY(10px); }
    100% { opacity: 1; transform: translateY(0); }
}

@keyframes pulseGlow {
    0%, 100% { transform: scale(1); opacity: 1; box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.4); }
    50% { transform: scale(1.1); opacity: 0.85; box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }
}

@keyframes ctaPulse {
    0% { box-shadow: 0 4px 14px rgba(37, 99, 235, 0.25); }
    50% { box-shadow: 0 6px 22px rgba(37, 99, 235, 0.45); }
    100% { box-shadow: 0 4px 14px rgba(37, 99, 235, 0.25); }
}

/* 3. Warm Eye-Comfort Animated Mesh Background (No Harsh Glare) */
.stApp, .stAppViewContainer {
    background: linear-gradient(-45deg, #EAE6DF, #E2E8F0, #E8ECEF, #EFECE6) !important;
    background-size: 300% 300% !important;
    animation: ambientBg 18s ease infinite !important;
}

/* 4. Frosted Alabaster Main Card */
.block-container {
    background: rgba(248, 246, 242, 0.9) !important;
    backdrop-filter: blur(16px) !important;
    -webkit-backdrop-filter: blur(16px) !important;
    padding: 2rem 2.75rem 2.25rem 2.75rem !important;
    border-radius: 20px !important;
    box-shadow: 
        0 20px 45px -15px rgba(30, 41, 59, 0.09),
        0 4px 12px -2px rgba(30, 41, 59, 0.04),
        inset 0 1px 0 0 rgba(255, 255, 255, 0.7) !important;
    max-width: 720px !important;
    margin-top: 1.6rem !important;
    margin-bottom: 1rem !important;
    border: 1px solid rgba(203, 213, 225, 0.65) !important;
    animation: cardReveal 0.65s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}

/* 5. Header & Title Styling */
.hero-wrap {
    text-align: center;
    animation: staggerFade 0.6s ease-out 0.1s both;
}

.custom-title {
    font-size: 2.15rem;
    font-weight: 800;
    margin-bottom: 0.45rem;
    letter-spacing: -0.03em;
    background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 55%, #334155 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}

.custom-subtitle {
    color: #57606A;
    font-size: 0.96rem;
    font-weight: 500;
    margin-bottom: 1.25rem;
    line-height: 1.5;
}

/* 6. Interactive GitHub Source Button (Fixes Link Underline) */
.github-wrapper {
    text-align: center;
    margin-bottom: 1.15rem;
    animation: staggerFade 0.6s ease-out 0.2s both;
}

a.github-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: #1E293B;
    color: #F8FAFC !important;
    padding: 0.52rem 1.25rem;
    border-radius: 9999px;
    font-size: 0.86rem;
    font-weight: 600;
    text-decoration: none !important;
    border: 1px solid #334155;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.12);
    transition: all 0.28s cubic-bezier(0.16, 1, 0.3, 1);
}

a.github-badge:hover {
    background: #2563EB;
    border-color: #3B82F6;
    color: #FFFFFF !important;
    text-decoration: none !important;
    transform: translateY(-2px) scale(1.02);
    box-shadow: 0 8px 20px rgba(37, 99, 235, 0.28);
}

a.github-badge svg {
    transition: transform 0.3s ease;
}

a.github-badge:hover svg {
    transform: rotate(-8deg) scale(1.1);
}

/* 7. Animated Tech Stack Pills */
.tech-stack {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
    margin-bottom: 1.5rem;
    animation: staggerFade 0.6s ease-out 0.3s both;
}

.tech-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background-color: #ECE9E2;
    color: #475569;
    padding: 5px 13px;
    border-radius: 9999px;
    font-size: 0.76rem;
    font-weight: 600;
    border: 1px solid #DCD7CD;
    transition: all 0.25s ease;
    cursor: default;
}

.tech-badge:hover {
    transform: translateY(-2px);
    background-color: #E2DFD7;
    border-color: #CBD5E1;
    color: #0F172A;
}

.tech-badge.highlight {
    background: linear-gradient(135deg, #EFF6FF 0%, #E0F2FE 100%);
    color: #0369A1;
    border: 1px solid #BAE6FD;
    box-shadow: 0 2px 8px rgba(2, 132, 199, 0.08);
}

.tech-badge.highlight:hover {
    box-shadow: 0 4px 12px rgba(2, 132, 199, 0.16);
    border-color: #7DD3FC;
}

.status-dot {
    width: 7px;
    height: 7px;
    background-color: #10B981;
    border-radius: 50%;
    display: inline-block;
    animation: pulseGlow 2s infinite;
}

/* 8. Uploader Label & Spinner */
label, [data-testid="stSpinner"] div {
    color: #334155 !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
}

/* 9. Animated File Uploader Dropzone */
[data-testid="stFileUploader"] {
    animation: staggerFade 0.6s ease-out 0.4s both;
}

[data-testid="stFileUploadDropzone"] {
    background-color: #F1EFEA !important;
    border: 2px dashed #C8C2B8 !important;
    border-radius: 14px !important;
    padding: 1.65rem 1.5rem !important;
    transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
}

[data-testid="stFileUploadDropzone"]:hover {
    border-color: #3B82F6 !important;
    background-color: #EBF3FE !important;
    transform: translateY(-2px);
    box-shadow: 0 10px 25px -5px rgba(59, 130, 246, 0.12) !important;
}

[data-testid="stFileUploadDropzone"] div, 
[data-testid="stFileUploadDropzone"] span {
    color: #57606A !important;
}

/* Upload Browse Button */
[data-testid="stFileUploadDropzone"] button {
    background-color: #F8F6F2 !important;
    color: #1E293B !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 8px !important;
    padding: 0.4rem 1rem !important;
    font-weight: 600 !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04) !important;
}

[data-testid="stFileUploadDropzone"] button:hover {
    background-color: #2563EB !important;
    color: #FFFFFF !important;
    border-color: #2563EB !important;
    transform: scale(1.03);
}

/* 10. Uploaded File Pill & Success Banner */
[data-testid="stUploadedFile"] {
    background-color: #F4F2ED !important;
    border: 1px solid #DFDBD2 !important;
    border-radius: 10px !important;
    animation: staggerFade 0.35s ease-out both;
}

[data-testid="stUploadedFile"] * {
    color: #1E293B !important;
}

[data-testid="stAlert"] {
    background-color: #ECFDF5 !important;
    border: 1px solid #A7F3D0 !important;
    border-radius: 10px !important;
    padding: 0.7rem 1rem !important;
    animation: staggerFade 0.4s ease-out both;
}

[data-testid="stAlert"] div {
    color: #065F46 !important;
    font-weight: 600 !important;
}

/* 11. Primary Animated Download Button */
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.55rem 1.35rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.01em !important;
    transition: all 0.28s cubic-bezier(0.16, 1, 0.3, 1) !important;
    animation: ctaPulse 2.5s infinite ease-in-out;
}

[data-testid="stDownloadButton"] button:hover {
    background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%) !important;
    transform: translateY(-2px) scale(1.02) !important;
    box-shadow: 0 10px 20px -5px rgba(37, 99, 235, 0.45) !important;
}

[data-testid="stDownloadButton"] button * {
    color: #FFFFFF !important;
}
</style>

<div class="hero-wrap">
    <div class="custom-title">Kannada ASCII to Unicode Converter</div>
    <div class="custom-subtitle">Upload your <code>.doc</code>, <code>.docx</code>, <code>.xls</code>, or <code>.xlsx</code> file to convert it to Unicode format.</div>
</div>

<div class="github-wrapper">
    <a href="https://github.com/NethenDatascientist/NUDI-to-Unicode-converter" target="_blank" class="github-badge">
        <svg height="16" width="16" viewBox="0 0 16 16" fill="currentColor">
            <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path>
        </svg>
        Code Link : GitHub Repository
    </a>
</div>

<div class="tech-stack">
    <span class="tech-badge">🐍 Python</span>
    <span class="tech-badge">⚡ Streamlit</span>
    <span class="tech-badge highlight">
        <span class="status-dot"></span>
        Logic: Sanka + Sanchaya Converting Logics
    </span>
</div>
""",
    unsafe_allow_html=True,
)


def get_secret(key, default=None):
    """Reads from Render Environment Variables first, then falls back to st.secrets."""
# ============================================================
# NON-BLOCKING VISITOR TRACKER & SHORT EMAIL ALERT
# ============================================================
def capture_visitor_context():
    headers = getattr(st.context, "headers", {}) or {}
    forwarded = headers.get("X-Forwarded-For", "")
    ip = (
        forwarded.split(",")[0].strip()
        if forwarded
        else (getattr(st.context, "ip_address", None) or "Unknown")
    )
    user_agent = headers.get("User-Agent", "Unknown Browser")
    os_platform = headers.get("Sec-Ch-Ua-Platform", "Unknown OS").strip('"')
    mobile = "?1" in headers.get("Sec-Ch-Ua-Mobile", "")
    device_type = "Mobile" if mobile else "Desktop"
    tz = getattr(st.context, "timezone", None) or "Unknown TZ"
    return ip, user_agent, f"{os_platform} ({device_type})", tz

def _send_alert_background(
    file_info, ip, user_agent, sys_platform, tz, sender_email, app_password, receiver_email
):
    try:
        loc_str = "Unknown Location"
        if ip and ip not in ["Unknown", "127.0.0.1", "::1"]:
            try:
                url = f"http://ip-api.com/json/{ip}?fields=status,city,regionName,isp"
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=2) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    if data.get("status") == "success":
                        loc_str = f"{data.get('city')}, {data.get('regionName')} ({data.get('isp')})"
            except Exception:
                pass

        ist_now = datetime.now(timezone(timedelta(hours=5, minutes=30))).strftime(
            "%d-%b-%Y %I:%M %p IST"
        )

        short_body = (
            f"File      : {file_info}\n"
            f"Time      : {ist_now}\n"
            f"System    : {sys_platform} | {tz}\n"
            f"Location  : {loc_str} [{ip}]\n"
            f"Browser   : {user_agent}\n"
        )

        print(f"\n--- 🔔 ACTIVITY ---\n{short_body}-------------------", flush=True)

        if sender_email and app_password:
            msg = EmailMessage()
            msg["Subject"] = f"🔤 File Converted: {file_info}"
            msg["From"] = sender_email
            msg["To"] = receiver_email
            msg.set_content(short_body)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=5) as server:
                server.login(sender_email, app_password)
                server.send_message(msg)
    except Exception as e:
        print(f"Email error: {e}", flush=True) # Added temporary print for debugging

def notify_file_upload(filename, file_size_bytes):
    try:
        size_kb = f"{file_size_bytes / 1024:.1f} KB"
        file_info = f"{filename} ({size_kb})"
        ip, user_agent, sys_platform, tz = capture_visitor_context()
        
        # Make sure these match exactly what you set in Render/Secrets
        sender_email = get_secret("NOTIFY_EMAIL")
        app_password = get_secret("NOTIFY_APP_PASSWORD")
        receiver_email = get_secret("RECEIVER_EMAIL", "dsdanethen@gmail.com")

        threading.Thread(
            target=_send_alert_background,
            args=(
                file_info, ip, user_agent, sys_platform, tz, 
                sender_email, app_password, receiver_email
            ),
            daemon=True,
        ).start()
    except Exception:
        pass
# ============================================================
# SANCHAYA NUDI/BARAHA ENGINE PORTED TO PYTHON
# ============================================================

VOWEL_MAPS = [
    ("A", "ಂ"), ("B", "ಃ"), ("CA", "ಅಂ"), ("CB", "ಅಃ"),
    ("C", "ಅ"), ("D", "ಆ"), ("E", "ಇ"), ("F", "ಈ"),
    ("G", "ಉ"), ("H", "ಊ"), ("IÄ", "ಋ"), ("IÆ", "ೠ"),
    ("J", "ಎ"), ("K", "ಏ"), ("L", "ಐ"), ("M", "ಒ"),
    ("N", "ಓ"), ("O", "ಔ"),
    ("x", "ಕ್ಷ"), ("GY", "ಜ್ಞ"),
]

CONSONANT_MAPS = [
    ("Pï", "ಕ್"), ("PÀ", "ಕ"), ("PÁ", "ಕಾ"), ("Q", "ಕಿ"), ("QÃ", "ಕೀ"),
    ("PÀÄ", "ಕು"), ("PÀÆ", "ಕೂ"), ("PÀÈ", "ಕೃ"), ("PÉ", "ಕೆ"), ("PÉÃ", "ಕೇ"),
    ("PÉÊ", "ಕೈ"), ("PÉÆ", "ಕೊ"), ("PÉÆÃ", "ಕೋ"), ("PË", "ಕೌ"),
    ("Sï", "ಖ್"), ("R", "ಖ"), ("SÁ", "ಖಾ"), ("T", "ಖಿ"), ("TÃ", "ಖೀ"),
    ("RÄ", "ಖು"), ("RÆ", "ಖೂ"), ("RÈ", "ಖೃ"), ("SÉ", "ಖೆ"), ("SÉÃ", "ಖೇ"),
    ("SÉÊ", "ಖೈ"), ("SÉÆ", "ಖೊ"), ("SÉÆÃ", "ಖೋ"), ("SË", "ಖೌ"),
    ("Uï", "ಗ್"), ("UÀ", "ಗ"), ("UÁ", "ಗಾ"), ("V", "ಗಿ"), ("VÃ", "ಗೀ"),
    ("UÀÄ", "ಗು"), ("UÀÆ", "ಗೂ"), ("UÀÈ", "ಗೃ"), ("UÉ", "ಗೆ"), ("UÉÃ", "ಗೇ"),
    ("UÉÊ", "ಗೈ"), ("UÉÆ", "ಗೊ"), ("UÉÆÃ", "ಗೋ"), ("UË", "ಗೌ"),
    ("Wï", "ಘ್"), ("WÀ", "ಘ"), ("WÁ", "ಘಾ"), ("X", "ಘಿ"), ("XÃ", "ಘೀ"),
    ("WÀÄ", "ಘು"), ("WÀÆ", "ಘೂ"), ("WÀÈ", "ಘೃ"), ("WÉ", "ಘೆ"), ("WÉÃ", "ಘೇ"),
    ("WÉÊ", "ಘೈ"), ("WÉÆ", "ಘೊ"), ("WÉÆÃ", "ಘೋ"), ("WË", "ಘೌ"),
    ("Yï", "ಙ್"), ("Y", "ಙ"),
    ("Zï", "ಚ್"), ("ZÀ", "ಚ"), ("ZÁ", "ಚಾ"), ("a", "ಚಿ"), ("aÃ", "ಚೀ"),
    ("ZÀÄ", "ಚು"), ("ZÀÆ", "ಚೂ"), ("ZÀÈ", "ಚೃ"), ("ZÉ", "ಚೆ"), ("ZÉÃ", "ಚೇ"),
    ("ZÉÊ", "ಚೈ"), ("ZÉÆ", "ಚೊ"), ("ZÉÆÃ", "ಚೋ"), ("ZË", "ಚೌ"),
    ("bï", "ಛ್"), ("bÀ", "ಛ"), ("bÁ", "ಛಾ"), ("c", "ಛಿ"), ("cÃ", "ಛೀ"),
    ("bÀÄ", "ಛು"), ("bÀÆ", "ಛೂ"), ("bÀÈ", "ಛೃ"), ("bÉ", "ಛೆ"), ("bÉÃ", "ಛೇ"),
    ("bÉÊ", "ಛೈ"), ("bÉÆ", "ಛೊ"), ("bÉÆÃ", "ಛೋ"), ("bË", "ಛೌ"),
    ("eï", "ಜ್"), ("d", "ಜ"), ("eÁ", "ಜಾ"), ("f", "ಜಿ"), ("fÃ", "ಜೀ"),
    ("dÄ", "ಜು"), ("dÆ", "ಜೂ"), ("dÈ", "ಜೃ"), ("eÉ", "ಜೆ"), ("eÉÃ", "ಜೇ"),
    ("eÉÊ", "ಜೈ"), ("eÉÆ", "ಜೊ"), ("eÉÆÃ", "ಜೋ"), ("eË", "ಜೌ"),
    ("gÀhiï", "ಝ್"), ("gÀhÄ", "ಝ"), ("gÀhiÁ", "ಝಾ"), ("jhÄ", "ಝಿ"), ("jhÄÃ", "ಝೀ"),
    ("gÀhÄÄ", "ಝು"), ("gÀhÄÆ", "ಝೂ"), ("gÀhÄÈ", "ಝೃ"), ("gÉhÄ", "ಝೆ"), ("gÉhÄÃ", "ಝೇ"),
    ("gÉhÄÊ", "ಝೈ"), ("gÉhÆ", "ಝೊ"), ("gÉhÆÃ", "ಝೋ"), ("gÀhiË", "ಝೌ"),
    ("kï", "ಞ್"), ("k", "ಞ"),
    ("mï", "ಟ್"), ("l", "ಟ"), ("mÁ", "ಟಾ"), ("n", "ಟಿ"), ("nÃ", "ಟೀ"),
    ("lÄ", "ಟು"), ("lÆ", "ಟೂ"), ("lÈ", "ಟೃ"), ("mÉ", "ಟೆ"), ("mÉÃ", "ಟೇ"),
    ("mÉÊ", "ಟೈ"), ("mÉÆ", "ಟೊ"), ("mÉÆÃ", "ಟೋ"), ("mË", "ಟೌ"),
    ("oï", "ಠ್"), ("oÀ", "ಠ"), ("oÁ", "ಠಾ"), ("p", "ಠಿ"), ("pÃ", "ಠೀ"),
    ("oÀÄ", "ಠು"), ("oÀÆ", "ಠೂ"), ("oÀÈ", "ಠೃ"), ("oÉ", "ಠೆ"), ("oÉÃ", "ಠೇ"),
    ("oÉÊ", "ಠೈ"), ("oÉÆ", "ಠೊ"), ("oÉÆÃ", "ಠೋ"), ("oË", "ಠೌ"),
    ("qï", "ಡ್"), ("qÀ", "ಡ"), ("qÁ", "ಡಾ"), ("r", "ಡಿ"), ("rÃ", "ಡೀ"),
    ("qÀÄ", "ಡು"), ("qÀÆ", "ಡೂ"), ("qÀÈ", "ಡೃ"), ("qÉ", "ಡೆ"), ("qÉÃ", "ಡೇ"),
    ("qÉÊ", "ಡೈ"), ("qÉÆ", "ಡೊ"), ("qÉÆÃ", "ಡೋ"), ("qË", "ಡೌ"),
    ("qsï", "ಢ್"), ("qsÀ", "ಢ"), ("qsÁ", "ಢಾ"), ("rü", "ಢಿ"), ("rüÃ", "ಢೀ"),
    ("qsÀÄ", "ಢು"), ("qsÀÆ", "ಢೂ"), ("qsÀÈ", "ಢೃ"), ("qsÉ", "ಢೆ"), ("qsÉÃ", "ಢೇ"),
    ("qsÉÊ", "ಢೈ"), ("qsÉÆ", "ಢೊ"), ("qsÉÆÃ", "ಢೋ"), ("qsË", "ಢೌ"),
    ("uï", "ಣ್"), ("t", "ಣ"), ("uÁ", "ಣಾ"), ("tÂ", "ಣಿ"), ("tÂÃ", "ಣೀ"),
    ("tÄ", "ಣು"), ("tÆ", "ಣೂ"), ("tÈ", "ಣೃ"), ("uÉ", "ಣೆ"), ("uÉÃ", "ಣೇ"),
    ("uÉÊ", "ಣೈ"), ("uÉÆ", "ಣೊ"), ("uÉÆÃ", "ಣೋ"), ("uË", "ಣೌ"),
    ("vï", "ತ್"), ("vÀ", "ತ"), ("vÁ", "ತಾ"), ("w", "ತಿ"), ("wÃ", "ತೀ"),
    ("vÀÄ", "ತು"), ("vÀÆ", "ತೂ"), ("vÀÈ", "ತೃ"), ("vÉ", "ತೆ"), ("vÉÃ", "ತೇ"),
    ("vÉÊ", "ತೈ"), ("vÉÆ", "ತೊ"), ("vÉÆÃ", "ತೋ"), ("vË", "ತೌ"),
    ("xï", "ಥ್"), ("xÀ", "ಥ"), ("xÁ", "ಥಾ"), ("y", "ಥಿ"), ("yÃ", "ಥೀ"),
    ("xÀÄ", "ಥು"), ("xÀÆ", "ಥೂ"), ("xÀÈ", "ಥೃ"), ("xÉ", "ಥೆ"), ("xÉÃ", "ಥೇ"),
    ("xÉÊ", "ಥೈ"), ("xÉÆ", "ಥೊ"), ("xÉÆÃ", "ಥೋ"), ("xË", "ಥೌ"),
    ("zï", "ದ್"), ("zÀ", "ದ"), ("zÁ", "ದಾ"), ("¢", "ದಿ"), ("¢Ã", "ದೀ"),
    ("zÀÄ", "ದು"), ("zÀÆ", "ದೂ"), ("zÀÈ", "ದೃ"), ("zÉ", "ದೆ"), ("zÉÃ", "ದೇ"),
    ("zÉÊ", "ದೈ"), ("zÉÆ", "ದೊ"), ("zÉÆÃ", "ದೋ"), ("zË", "ದೌ"),
    ("zsï", "ಧ್"), ("zsÀ", "ಧ"), ("zsÁ", "ಧಾ"), ("¢ü", "ಧಿ"), ("¢üÃ", "ಧೀ"),
    ("zsÀÄ", "ಧು"), ("zsÀÆ", "ಧೂ"), ("zsÀÈ", "ಧೃ"), ("zsÉ", "ಧೆ"), ("zsÉÃ", "ಧೇ"),
    ("zsÉÊ", "ಧೈ"), ("zsÉÆ", "ಧೊ"), ("zsÉÆÃ", "ಧೋ"), ("zsË", "ಧೌ"),
    ("£ï", "ನ್"), ("£À", "ನ"), ("£Á", "ನಾ"), ("¤", "ನಿ"), ("¤Ã", "ನೀ"),
    ("£ÀÄ", "ನು"), ("£ÀÆ", "ನೂ"), ("£ÀÈ", "ನೃ"), ("£É", "ನೆ"), ("£ÉÃ", "ನೇ"),
    ("£ÉÊ", "ನೈ"), ("£ÉÆ", "ನೊ"), ("£ÉÆÃ", "ನೋ"), ("£Ë", "ನೌ"),
    ("¥ï", "ಪ್"), ("¥À", "ಪ"), ("¥Á", "ಪಾ"), ("¦", "ಪಿ"), ("¦Ã", "ಪೀ"),
    ("¥ÀÄ", "ಪು"), ("¥ÀÅ", "ಪು"), ("¥ÀÆ", "ಪೂ"), ("¥ÀÇ", "ಪೂ"), ("¥ÀÈ", "ಪೃ"),
    ("¥É", "ಪೆ"), ("¥ÉÃ", "ಪೇ"), ("¥ÉÊ", "ಪೈ"), ("¥ÉÆ", "ಪೊ"), ("¥ÉÇ", "ಪೊ"),
    ("¥ÉÆÃ", "ಪೋ"), ("¥ÉÇÃ", "ಪೋ"), ("¥Ë", "ಪೌ"),
    ("¥sï", "ಫ್"), ("¥sÀ", "ಫ"), ("¥sÁ", "ಫಾ"), ("¦ü", "ಫಿ"), ("¦üÃ", "ಫೀ"),
    ("¥sÀÄ", "ಫು"), ("¥sÀÅ", "ಫು"), ("¥sÀÆ", "ಫೂ"), ("¥sÀÇ", "ಫೂ"), ("¥sÀÈ", "ಫೃ"),
    ("¥sÉ", "ಫೆ"), ("¥sÉÃ", "ಫೇ"), ("¥sÉÊ", "ಫೈ"), ("¥sÉÆ", "ಫೊ"), ("¥sÉÇ", "ಫೊ"),
    ("¥sÉÆÃ", "ಫೋ"), ("¥sÉÇÃ", "ಫೋ"), ("¥sË", "ಫೌ"),
    ("¨ï", "ಬ್"), ("§", "ಬ"), ("¨Á", "ಬಾ"), ("©", "ಬಿ"), ("©Ã", "ಬೀ"),
    ("§Ä", "ಬು"), ("§Æ", "ಬೂ"), ("§È", "ಬೃ"), ("¨É", "ಬೆ"), ("¨ÉÃ", "ಬೇ"),
    ("¨ÉÊ", "ಬೈ"), ("¨ÉÆ", "ಬೊ"), ("¨ÉÆÃ", "ಬೋ"), ("¨Ë", "ಬೌ"),
    ("¨sï", "ಭ್"), ("¨sÀ", "ಭ"), ("¨sÁ", "ಭಾ"), ("©ü", "ಭಿ"), ("©üÃ", "ಭೀ"),
    ("¨sÀÄ", "ಭು"), ("¨sÀÆ", "ಭೂ"), ("¨sÀÈ", "ಭೃ"), ("¨sÉ", "ಭೆ"), ("¨sÉÃ", "ಭೇ"),
    ("¨sÉÊ", "ಭೈ"), ("¨sÉÆ", "ಭೊ"), ("¨sÉÆÃ", "ಭೋ"), ("¨sË", "ಭೌ"),
    ("ªÀiï", "ಮ್"), ("ªÀÄ", "ಮ"), ("ªÀiÁ", "ಮಾ"), ("«Ä", "ಮಿ"), ("«ÄÃ", "ಮೀ"),
    ("ªÀÄÄ", "ಮು"), ("ªÀÄÆ", "ಮೂ"), ("ªÀÄÈ", "ಮೃ"), ("ªÉÄ", "ಮೆ"), ("ªÉÄÃ", "ಮೇ"),
    ("ªÉÄÊ", "ಮೈ"), ("ªÉÆ", "ಮೊ"), ("ªÉÆÃ", "ಮೋ"), ("ªÀiË", "ಮೌ"),
    ("AiÀiï", "ಯ್"), ("AiÀÄ", "ಯ"), ("AiÀiÁ", "ಯಾ"), ("¬Ä", "ಯಿ"), ("¬ÄÃ", "ಯೀ"),
    ("AiÀÄÄ", "ಯು"), ("AiÀÄÆ", "ಯೂ"), ("AiÀÄÈ", "ಯೃ"), ("AiÉÄ", "ಯೆ"), ("AiÉÄÃ", "ಯೇ"),
    ("AiÉÄÊ", "ಯೈ"), ("AiÉÆ", "ಯೊ"), ("AiÉÆÃ", "ಯೋ"), ("AiÀiË", "ಯೌ"),
    ("gï", "ರ್"), ("gÀ", "ರ"), ("gÁ", "ರಾ"), ("j", "ರಿ"), ("jÃ", "ರೀ"),
    ("gÀÄ", "ರು"), ("gÀÆ", "ರೂ"), ("gÀÈ", "ರೃ"), ("gÉ", "ರೆ"), ("gÉÃ", "ರೇ"),
    ("gÉÊ", "ರೈ"), ("gÉÆ", "ರೊ"), ("gÉÆÃ", "ರೋ"), ("gË", "ರೌ"),
    ("¯ï", "ಲ್"), ("¯i", "ಲ್"), ("®", "ಲ"), ("¯Á", "ಲಾ"), ("°", "ಲಿ"), ("°Ã", "ಲೀ"),
    ("®Ä", "ಲು"), ("®Æ", "ಲೂ"), ("®È", "ಲೃ"), ("¯É", "ಲೆ"), ("¯ÉÃ", "ಲೇ"),
    ("¯ÉÊ", "ಲೈ"), ("¯ÉÆ", "ಲೊ"), ("¯ÉÆÃ", "ಲೋ"), ("¯Ë", "ಲೌ"),
    ("ªï", "ವ್"), ("ªÀ", "ವ"), ("ªÁ", "ವಾ"), ("«", "ವಿ"), ("«Ã", "ವೀ"),
    ("ªÀÅ", "ವು"), ("ªÀÇ", "ವೂ"), ("ªÀÈ", "ವೃ"), ("ªÉ", "ವೆ"), ("ªÉÃ", "ವೇ"),
    ("ªÉÊ", "ವೈ"), ("ªÉÇ", "ವೊ"), ("ªÉÇÃ", "ವೋ"), ("ªË", "ವೌ"),
    ("±ï", "ಶ್"), ("±À", "ಶ"), ("±Á", "ಶಾ"), ("²", "ಶಿ"), ("²Ã", "ಶೀ"),
    ("±ÀÄ", "ಶು"), ("±ÀÆ", "ಶೂ"), ("±ÀÈ", "ಶೃ"), ("±É", "ಶೆ"), ("±ÉÃ", "ಶೇ"),
    ("±ÉÊ", "ಶೈ"), ("±ÉÆ", "ಶೊ"), ("±ÉÆÃ", "ಶೋ"), ("±Ë", "ಶೌ"),
    ("μï", "ಷ್"), ("μÀ", "ಷ"), ("μÁ", "ಷಾ"), ("¶", "ಷಿ"), ("¶Ã", "ಷೀ"),
    ("μÀÄ", "ಷು"), ("μÀÆ", "ಷೂ"), ("μÀÈ", "ಷೃ"), ("μÉ", "ಷೆ"), ("μÉÃ", "ಷೇ"),
    ("μÉÊ", "ಷೈ"), ("μÉÆ", "ಷೊ"), ("μÉÆÃ", "ಷೋ"), ("μË", "ಷೌ"),
    ("¸ï", "ಸ್"), ("¸À", "ಸ"), ("¸Á", "ಸಾ"), ("¹", "ಸಿ"), ("¹Ã", "ಸೀ"),
    ("¸ÀÄ", "ಸು"), ("¸ÀÆ", "ಸೂ"), ("¸ÀÈ", "ಸೃ"), ("¸É", "ಸೆ"), ("¸ÉÃ", "ಸೇ"),
    ("¸ÉÊ", "ಸೈ"), ("¸ÉÆ", "ಸೊ"), ("¸ÉÆÃ", "ಸೋ"), ("¸Ë", "ಸೌ"),
    ("ºï", "ಹ್"), ("ºÀ", "ಹ"), ("ºÁ", "ಹಾ"), ("»", "ಹಿ"), ("»Ã", "ಹೀ"),
    ("ºÀÄ", "ಹು"), ("ºÀÆ", "ಹೂ"), ("ºÀÈ", "ಹೃ"), ("ºÉ", "ಹೆ"), ("ºÉÃ", "ಹೇ"),
    ("ºÉÊ", "ಹೈ"), ("ºÉÆ", "ಹೊ"), ("ºÉÆÃ", "ಹೋ"), ("ºË", "ಹೌ"),
    ("¼ï", "ಳ್"), ("¼À", "ಳ"), ("¼Á", "ಳಾ"), ("½", "ಳಿ"), ("½Ã", "ಳೀ"),
    ("¼ÀÄ", "ಳು"), ("¼ÀÆ", "ಳೂ"), ("¼ÀÈ", "ಳೃ"), ("¼É", "ಳೆ"), ("¼ÉÃ", "ಳೇ"),
    ("¼ÉÊ", "ಳೈ"), ("¼ÉÆ", "ಳೊ"), ("¼ÉÆÃ", "ಳೋ"), ("¼Ë", "ಳೌ"),
]

A2U_MAP = {}
for k, v in VOWEL_MAPS:
    A2U_MAP[k] = v
for k, v in CONSONANT_MAPS:
    A2U_MAP[k] = v

# Sort keys longest-first and compile into a single master regex for 10x faster replacement
A2U_KEYS = [k for k in sorted(A2U_MAP.keys(), key=len, reverse=True) if k and k != "ï"]
A2U_MASTER_RE = re.compile("|".join(re.escape(k) for k in A2U_KEYS))

VATTAKSHARA_MAP = {
    "Ì": "ಕ್", "Í": "ಖ್", "Î": "ಗ್", "Ï": "ಘ್", "Ð": "ಙ್",
    "Ñ": "ಚ್", "Ò": "ಛ್", "Ó": "ಜ್", "Ô": "ಝ್", "Õ": "ಞ್",
    "Ö": "ಟ್", "×": "ಠ್", "Ø": "ಡ್", "Ù": "ಢ್", "Ú": "ಣ್",
    "Û": "ತ್", "Ü": "ಥ್", "Ý": "ದ್", "Þ": "ಧ್", "ß": "ನ್",
    "à": "ಪ್", "á": "ಫ್", "â": "ಬ್", "ã": "ಭ್", "ä": "ಮ್",
    "å": "ಯ್", "æ": "ರ್", "è": "ಲ್", "é": "ವ್", "ê": "ಶ್",
    "ë": "ಷ್", "ì": "ಸ್", "í": "ಹ್", "î": "ಳ್",
    "ç": "ರ್", "ö": "ಯ್",
}

VATT_MARKER = "\u0001"

OTHER_MAP = {
    "ø": "ೃ",
    "ñ": "ೄ",
    "„": "ಽ",
    "ó": "಼",
    "ð": "ರ",
    "ï": "್",
}

ASCII_VATTAKSHARA = "ÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîö"
DEP_VOWELS = "್ಾಿೀುೂೃೆೇೈೊೋೌ"

# Pre-compiled Regexes for Sanchaya pipeline stages
VATT_REORDER_3_RE = re.compile(
    r"([ೆೇೊ])([ÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæèéêëìíî])([ÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæèéêëìíî])([ÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæèéêëìíî])"
)
VATT_REORDER_2_RE = re.compile(
    r"([ೆೇೊ])([ÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæèéêëìíî])([ÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæèéêëìíî])"
)
VATT_REORDER_1_RE = re.compile(r"([ೆೇೊ])([ÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæèéêëìíî])")
DUP_E_RE = re.compile(r"É{2,}")

ASPIRATE_VOWEL_RUN = r"É(?:ÆÃ|Æ|Ê|Ã)?|À(?:Ä|Å|Æ|Ç|È)?|[ÁË]"
ASPIRATE_REORDER_RE = re.compile(rf"([¥¨qz])({ASPIRATE_VOWEL_RUN})s")

VATTU_VOWEL_RUN = (
    r"É(?:ÆÃ|Æ|Ê|Ã)?|À(?:Ä|Å|Æ|Ç|È)?|Â(?:Ã)?|[ÁË]|"
    r"rü|¢ü|¦ü|©ü|jhÄ|«Ä|Q|T|V|X|a|c|f|n|p|r|w|y|¢|¦|©|j|°|«|²|¶|¹|»|½"
)
VATTU_BEFORE_A_RE = re.compile(rf"([{ASCII_VATTAKSHARA}])({VATTU_VOWEL_RUN})")

REPH_RE = re.compile(rf"([ಕ-ಹೞ])([{DEP_VOWELS}]?)ð")
CONJUNCT_RE = re.compile(rf"([ಕ-ಹೞ])([ಾಿೀುೂೃೄೆೇೈೊೋೌ]?){VATT_MARKER}([ಕ-ಹೞ])")
MULTI_HALANT_RE = re.compile(r"([ಕ-ಹ])(್){2,}")

ORPHAN_U_RE = re.compile(r"([ಕ-ಹೞ])[ÄÅ]")
ORPHAN_UU_RE = re.compile(r"([ಕ-ಹೞ])[ÆÇ]")
ORPHAN_R_RE = re.compile(r"([ಕ-ಹೞ])È")

DEERGA_RE = re.compile(r"([ಿೆೇೊ])Ã")
UNI_DEERGA_MAP = {"ಿ": "ೀ", "ೆ": "ೇ", "ೊ": "ೋ"}

EMAIL_RE = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
URL_RE = re.compile(r"https?://\S+|www\.\S+")
SPLIT_RE = re.compile(r"(\s+)")
ACRONYM_OR_NUM_RE = re.compile(r"^[A-Z0-9\-\./]+$|^[0-9\-\./]+$")
ASCII_CHECK_RE = re.compile(r"[\x21-\x7E\x80-\xFFµ]")

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


def sanchaya_nudi_word_to_unicode(word):
    """Executes Sanchaya's exact multi-stage Nudi-to-Unicode pipeline on a single token."""
    # Normalize Windows-1252 micro sign to Greek small mu
    converted = word.replace("µ", "μ")

    # 1. Pre-reorder vowel signs before vattakshara clusters
    converted = VATT_REORDER_3_RE.sub(r"\2\3\4\1", converted)
    converted = VATT_REORDER_2_RE.sub(r"\2\3\1", converted)
    converted = VATT_REORDER_1_RE.sub(r"\2\1", converted)

    # 2. Collapse duplicate 'É' markers
    converted = DUP_E_RE.sub("É", converted)

    # 3. Reorder aspirate 's' marker typed after vowel back to before vowel
    converted = ASPIRATE_REORDER_RE.sub(r"\1s\2", converted)

    # 4. Reorder vattakshara typed before vowel marker back to canonical order
    converted = VATTU_BEFORE_A_RE.sub(r"\2\1", converted)

    # 5. Core longest-first dictionary replacement
    converted = A2U_MASTER_RE.sub(lambda m: A2U_MAP[m.group(0)], converted)

    # 6. Strip orphaned 'À' markers (fixes the 'ಭೂಸ್ವಾಪÀ್ ೕರ' bug)
    converted = converted.replace("À", "")

    # 7. Move 'ð' (reph / ಅರ್ಕಾವೊತ್ತು) to precede the syllable
    converted = REPH_RE.sub(r"ರ್\1\2", converted)

    # 8. Replace vattakshara & special characters
    for k, v in VATTAKSHARA_MAP.items():
        if k in converted:
            converted = converted.replace(k, VATT_MARKER + v[:-1])
    for k, v in OTHER_MAP.items():
        if k in converted:
            converted = converted.replace(k, v)

    # 9. Resolve conjunct chains (vattakshara reordering)
    while CONJUNCT_RE.search(converted):
        converted = CONJUNCT_RE.sub(r"\1್\3\2", converted)
    converted = converted.replace(VATT_MARKER, "್")
    converted = MULTI_HALANT_RE.sub(r"\1\2", converted)

    # 10. Attach orphan u/uu/vocalic-r vowels after conjuncts
    converted = ORPHAN_U_RE.sub(r"\1ು", converted)
    converted = ORPHAN_UU_RE.sub(r"\1ೂ", converted)
    converted = ORPHAN_R_RE.sub(r"\1ೃ", converted)

    # 11. Anuswara & Visarga
    converted = converted.replace("A", "ಂ").replace("B", "ಃ")

    # 12. Post-process leftover vowel markers
    converted = (
        converted.replace("ÉÆÃ", "ೋ")
        .replace("ÉÆ", "ೊ")
        .replace("ÉÊ", "ೈ")
        .replace("ÉÃ", "ೇ")
        .replace("É", "ೆ")
        .replace("Á", "ಾ")
        .replace("Ë", "ೌ")
    )

    # 13. Deerga (vowel lengthener) handling
    converted = DEERGA_RE.sub(
        lambda m: UNI_DEERGA_MAP.get(m.group(1), m.group(1)), converted
    )
    converted = converted.replace("ುÆ", "ೂ")

    return converted


@functools.lru_cache(maxsize=50000)
def convert_single_token(token):
    """Preserves English words, numbers, and acronyms; converts Nudi tokens via Sanchaya engine."""
    stripped = token.strip()
    if not stripped or not ASCII_CHECK_RE.search(stripped):
        return token

    cleaned_token = stripped.strip(".,;:!?()[]{}'\"-").lower()
    if ACRONYM_OR_NUM_RE.match(stripped) or cleaned_token in COMMON_ENGLISH_WORDS:
        return token

    # If a token has slashes (like government file numbers: ಸನೀಇ/132/ಭೂಸ್ವಾ/2023),
    # split by slash and convert each segment so numbers stay English and text becomes Kannada
    if "/" in token and not token.startswith("http"):
        sub_parts = token.split("/")
        return "/".join(convert_single_token(p) for p in sub_parts)

    try:
        return sanchaya_nudi_word_to_unicode(token)
    except Exception:
        return token


@functools.lru_cache(maxsize=20000)
def transliterate_text_safely(text):
    if not isinstance(text, str) or not text.strip():
        return text

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


# ============================================================
# DOCUMENT PROCESSING (LIBREOFFICE + DOCX + OPENPYXL)
# ============================================================
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
                if val_str and not val_str.startswith("="):
                    cell.value = transliterate_text_safely(val)

    wb.save(output_path)


# ============================================================
# STREAMLIT UI
# ============================================================
uploaded_file = st.file_uploader(
    "ಫೈಲ್ ಅನ್ನು ಆಯ್ಕೆ ಮಾಡಿ (Choose file)", type=["doc", "docx", "xls", "xlsx"]
)

if uploaded_file is not None:
    filename = uploaded_file.name

    if st.session_state.get("last_tracked_file") != filename:
        notify_file_upload(filename, uploaded_file.size)
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
