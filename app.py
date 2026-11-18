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
from fastapi import FastAPI, File, Form, UploadFile, Request, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, FileResponse

app = FastAPI(title="MI - Kannada ASCII to Unicode Converter")


def get_secret(key, default=None):
    """Reads from Environment Variables."""
    val = os.environ.get(key)
    return val if val else default


# ============================================================
# NON-BLOCKING VISITOR TRACKER & SHORT EMAIL ALERT
# ============================================================
def capture_visitor_context(request: Request, client_tz: str = "Asia/Kolkata"):
    headers = request.headers
    forwarded = headers.get("x-forwarded-for", "")
    ip = (
        forwarded.split(",")[0].strip()
        if forwarded
        else (request.client.host if request.client else "Unknown")
    )
    user_agent = headers.get("user-agent", "Unknown Browser")
    os_platform = headers.get("sec-ch-ua-platform", "Unknown OS").strip('"')
    mobile = "?1" in headers.get("sec-ch-ua-mobile", "")
    device_type = "Mobile" if mobile else "Desktop"
    tz = client_tz or "Unknown TZ"
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
        print(f"Email error: {e}", flush=True)


def notify_file_upload(request: Request, filename: str, file_size_bytes: int, client_tz: str):
    try:
        size_kb = f"{file_size_bytes / 1024:.1f} KB"
        file_info = f"{filename} ({size_kb})"
        ip, user_agent, sys_platform, tz = capture_visitor_context(request, client_tz)

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
    converted = word.replace("µ", "μ")

    converted = VATT_REORDER_3_RE.sub(r"\2\3\4\1", converted)
    converted = VATT_REORDER_2_RE.sub(r"\2\3\1", converted)
    converted = VATT_REORDER_1_RE.sub(r"\2\1", converted)

    converted = DUP_E_RE.sub("É", converted)
    converted = ASPIRATE_REORDER_RE.sub(r"\1s\2", converted)
    converted = VATTU_BEFORE_A_RE.sub(r"\2\1", converted)
    converted = A2U_MASTER_RE.sub(lambda m: A2U_MAP[m.group(0)], converted)
    converted = converted.replace("À", "")
    converted = REPH_RE.sub(r"ರ್\1\2", converted)

    for k, v in VATTAKSHARA_MAP.items():
        if k in converted:
            converted = converted.replace(k, VATT_MARKER + v[:-1])
    for k, v in OTHER_MAP.items():
        if k in converted:
            converted = converted.replace(k, v)

    while CONJUNCT_RE.search(converted):
        converted = CONJUNCT_RE.sub(r"\1್\3\2", converted)
    converted = converted.replace(VATT_MARKER, "್")
    converted = MULTI_HALANT_RE.sub(r"\1\2", converted)

    converted = ORPHAN_U_RE.sub(r"\1ು", converted)
    converted = ORPHAN_UU_RE.sub(r"\1ೂ", converted)
    converted = ORPHAN_R_RE.sub(r"\1ೃ", converted)

    converted = converted.replace("A", "ಂ").replace("B", "ಃ")

    converted = (
        converted.replace("ÉÆÃ", "ೋ")
        .replace("ÉÆ", "ೊ")
        .replace("ÉÊ", "ೈ")
        .replace("ÉÃ", "ೇ")
        .replace("É", "ೆ")
        .replace("Á", "ಾ")
        .replace("Ë", "ೌ")
    )

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
    except Exception as e:
        print(f"LibreOffice conversion error: {e}", flush=True)
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
# FASTAPI ROUTES & ANIMATED WEB INTERFACE
# ============================================================
@app.post("/convert")
async def convert_document(
    request: Request,
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    tz: str = Form("Asia/Kolkata"),
):
    filename = file.filename
    base_name, file_ext = os.path.splitext(filename)
    file_ext = file_ext.lower()

    if file_ext not in [".doc", ".docx", ".xls", ".xlsx"]:
        raise HTTPException(status_code=400, detail="Unsupported file format.")

    tmpdir = tempfile.mkdtemp()
    input_path = os.path.join(tmpdir, filename)
    content = await file.read()

    with open(input_path, "wb") as f:
        f.write(content)

    notify_file_upload(request, filename, len(content), tz)

    output_ext = ".docx" if file_ext in [".doc", ".docx"] else ".xlsx"
    download_filename = f"kannada_unicode_{base_name}{output_ext}"
    final_output_path = os.path.join(tmpdir, download_filename)

    if file_ext == ".doc":
        converted_path = convert_with_libreoffice(input_path, "docx")
        if converted_path:
            process_word_document(converted_path, final_output_path)
    elif file_ext == ".docx":
        process_word_document(input_path, final_output_path)
    elif file_ext == ".xls":
        converted_path = convert_with_libreoffice(input_path, "xlsx")
        if converted_path:
            process_excel_document(converted_path, final_output_path)
    elif file_ext == ".xlsx":
        process_excel_document(input_path, final_output_path)

    if not os.path.exists(final_output_path):
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise HTTPException(status_code=500, detail="Conversion failed or timed out.")

    background_tasks.add_task(shutil.rmtree, tmpdir, ignore_errors=True)
    return FileResponse(
        final_output_path,
        filename=download_filename,
        media_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if output_ext == ".docx"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


@app.get("/", response_class=HTMLResponse)
async def serve_ui():
    return """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MI - Kannada ASCII to Unicode Converter</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            transition: background-color 0.4s ease, color 0.4s ease;
        }

        /* 1. Warm Paper Theme (Default Eye-Care) */
        body.theme-warm { background-color: #EFECE6; color: #2D2723; }
        body.theme-warm .app-card { background: rgba(253, 251, 247, 0.94); border-color: #E2DDD3; box-shadow: 0 25px 50px -12px rgba(120, 110, 95, 0.14); }
        body.theme-warm .dropzone { background: #F5F2EB; border-color: #D5CFC2; }
        body.theme-warm .subtext { color: #6E655F; }
        body.theme-warm .pill { background: #E8E3D9; color: #4A433E; border-color: #D5CFC2; }

        /* 2. Soft Slate Theme */
        body.theme-slate { background-color: #E2E8F0; color: #0F172A; }
        body.theme-slate .app-card { background: rgba(248, 250, 252, 0.94); border-color: #CBD5E1; box-shadow: 0 25px 50px -12px rgba(15, 23, 42, 0.09); }
        body.theme-slate .dropzone { background: #F1F5F9; border-color: #CBD5E1; }
        body.theme-slate .subtext { color: #475569; }
        body.theme-slate .pill { background: #E2E8F0; color: #334155; border-color: #CBD5E1; }

        /* 3. Midnight Dark Theme */
        body.theme-dark { background-color: #0B1120; color: #F8FAFC; }
        body.theme-dark .app-card { background: rgba(30, 41, 59, 0.92); border-color: #334155; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.55); }
        body.theme-dark .dropzone { background: #0F172A; border-color: #475569; }
        body.theme-dark .subtext { color: #94A3B8; }
        body.theme-dark .pill { background: #0F172A; color: #CBD5E1; border-color: #334155; }

        @keyframes floatSlow {
            0%, 100% { transform: translateY(0px); }
            50% { transform: translateY(-6px); }
        }
        @keyframes cardEnter {
            from { opacity: 0; transform: translateY(22px) scale(0.98); }
            to { opacity: 1; transform: translateY(0) scale(1); }
        }
        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .animate-float { animation: floatSlow 3.2s ease-in-out infinite; }
        .animate-enter { animation: cardEnter 0.65s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
        .stagger-1 { animation: fadeUp 0.5s ease-out 0.1s both; }
        .stagger-2 { animation: fadeUp 0.5s ease-out 0.2s both; }
        .stagger-3 { animation: fadeUp 0.5s ease-out 0.3s both; }

        .dropzone { transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1); }
        .dropzone.dragover {
            border-color: #4F46E5 !important;
            transform: scale(1.015);
            box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.15);
        }
    </style>
</head>
<body class="theme-warm h-screen w-screen flex flex-col items-center justify-center p-4 overflow-hidden select-none">

    <!-- Top-Right Eye-Care Theme Switcher -->
    <div class="fixed top-4 right-4 flex items-center gap-1.5 p-1.5 rounded-full bg-black/5 backdrop-blur-md border border-black/10 z-50">
        <button onclick="setTheme('theme-warm')" class="px-3 py-1 rounded-full text-xs font-semibold transition hover:scale-105 shadow-sm" style="background:#FDFBF7; color:#2D2723;">🌅 Warm Paper</button>
        <button onclick="setTheme('theme-slate')" class="px-3 py-1 rounded-full text-xs font-semibold transition hover:scale-105 shadow-sm" style="background:#F1F5F9; color:#0F172A;">☁️ Soft Slate</button>
        <button onclick="setTheme('theme-dark')" class="px-3 py-1 rounded-full text-xs font-semibold transition hover:scale-105 shadow-sm" style="background:#1E293B; color:#F8FAFC;">🌙 Midnight</button>
    </div>

    <!-- Main Centered Card -->
    <div class="app-card animate-enter w-full max-w-2xl rounded-3xl border p-7 md:p-9 backdrop-blur-xl transition-all duration-300">
        
        <!-- Header Section -->
        <div class="text-center stagger-1">
            <h1 class="text-2xl md:text-3xl font-extrabold tracking-tight mb-2">
                MI - Kannada ASCII to Unicode Converter
            </h1>
            <p class="subtext text-sm md:text-base font-medium mb-5">
                Upload your <code class="px-1.5 py-0.5 rounded bg-black/5 text-xs font-bold">.doc</code>, <code class="px-1.5 py-0.5 rounded bg-black/5 text-xs font-bold">.docx</code>, <code class="px-1.5 py-0.5 rounded bg-black/5 text-xs font-bold">.xls</code>, or <code class="px-1.5 py-0.5 rounded bg-black/5 text-xs font-bold">.xlsx</code> file to convert it to Unicode format.
            </p>
        </div>

        <!-- GitHub Link & Tech Stack Badges -->
        <div class="flex flex-col items-center gap-3 mb-6 stagger-2">
            <a href="https://github.com/NethenDatascientist/NUDI-to-Unicode-converter" target="_blank"
               class="inline-flex items-center gap-2 px-5 py-2 rounded-full text-xs md:text-sm font-semibold bg-indigo-600 text-white hover:bg-indigo-500 transition-all duration-200 hover:-translate-y-0.5 shadow-md shadow-indigo-600/25">
                <svg class="w-4 h-4" fill="currentColor" viewBox="0 0 16 16"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"/></svg>
                Code Link : GitHub Repository
            </a>

            <div class="flex flex-wrap items-center justify-center gap-2">
                <span class="pill px-3 py-1 rounded-full text-xs font-semibold border">🐍 Python</span>
                <span class="pill px-3 py-1 rounded-full text-xs font-semibold border">⚡ FastAPI</span>
                <span class="pill px-3 py-1 rounded-full text-xs font-semibold border flex items-center gap-1.5">
                    <span class="relative flex h-2 w-2">
                        <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                        <span class="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
                    </span>
                    Logic: Sanka + Sanchaya Converting Logics
                </span>
            </div>
        </div>

        <!-- Interactive Upload & Conversion Area -->
        <div class="stagger-3">
            <label id="dropzone" class="dropzone relative flex flex-col items-center justify-center w-full h-48 border-2 border-dashed rounded-2xl cursor-pointer hover:border-indigo-500 group">
                <input type="file" id="fileInput" class="hidden" accept=".doc,.docx,.xls,.xlsx" />
                
                <!-- 1. Idle Upload State -->
                <div id="idleState" class="flex flex-col items-center text-center p-4">
                    <div class="animate-float w-12 h-12 mb-3 rounded-2xl bg-indigo-500/10 text-indigo-600 flex items-center justify-center group-hover:scale-110 transition">
                        <svg class="w-6 h-6" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
                        </svg>
                    </div>
                    <p class="font-bold text-sm md:text-base mb-1">ಫೈಲ್ ಅನ್ನು ಆಯ್ಕೆ ಮಾಡಿ (Click or Drag & Drop file)</p>
                    <p class="subtext text-xs">Supports .DOC, .DOCX, .XLS, .XLSX • Max 200MB</p>
                </div>

                <!-- 2. Animated Processing State -->
                <div id="loadingState" class="hidden flex-col items-center text-center w-full px-8">
                    <div class="w-10 h-10 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mb-3"></div>
                    <p id="statusText" class="font-bold text-sm mb-2">Processing and translating text... Please wait.</p>
                    <div class="w-full max-w-xs bg-black/10 h-2 rounded-full overflow-hidden">
                        <div id="progressBar" class="bg-indigo-600 h-full w-1/4 rounded-full transition-all duration-500"></div>
                    </div>
                </div>

                <!-- 3. Completed State -->
                <div id="successState" class="hidden flex-col items-center text-center p-4">
                    <div class="w-11 h-11 mb-2 rounded-full bg-emerald-500/15 text-emerald-600 flex items-center justify-center">
                        <svg class="w-6 h-6" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/>
                        </svg>
                    </div>
                    <p class="font-bold text-base text-emerald-600 mb-0.5">✅ Processing complete!</p>
                    <p id="convertedFileName" class="subtext text-xs mb-3"></p>
                    <div class="flex items-center gap-2.5">
                        <button type="button" id="downloadAgainBtn"
                                class="px-4 py-2 rounded-xl bg-indigo-600 text-white text-xs font-semibold shadow hover:bg-indigo-500 transition hover:-translate-y-0.5">
                            📥 Download Converted File
                        </button>
                        <button type="button" id="resetBtn"
                                class="pill px-3.5 py-2 rounded-xl text-xs font-semibold border hover:opacity-80 transition">
                            🔄 Convert Another
                        </button>
                    </div>
                </div>
            </label>
        </div>
    </div>

    <script>
        function setTheme(theme) {
            document.body.className = theme + " h-screen w-screen flex flex-col items-center justify-center p-4 overflow-hidden select-none";
            localStorage.setItem("preferredTheme", theme);
        }
        const savedTheme = localStorage.getItem("preferredTheme");
        if (savedTheme) setTheme(savedTheme);

        const dropzone = document.getElementById('dropzone');
        const fileInput = document.getElementById('fileInput');
        const idleState = document.getElementById('idleState');
        const loadingState = document.getElementById('loadingState');
        const successState = document.getElementById('successState');
        const progressBar = document.getElementById('progressBar');
        const statusText = document.getElementById('statusText');
        const downloadAgainBtn = document.getElementById('downloadAgainBtn');
        const resetBtn = document.getElementById('resetBtn');

        let lastBlobUrl = null;
        let lastOutName = null;

        ['dragenter', 'dragover'].forEach(evt => {
            dropzone.addEventListener(evt, (e) => {
                e.preventDefault();
                dropzone.classList.add('dragover');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            dropzone.addEventListener(evt, (e) => {
                e.preventDefault();
                dropzone.classList.remove('dragover');
            });
        });

        dropzone.addEventListener('drop', (e) => {
            if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) handleFile(e.target.files[0]);
        });

        downloadAgainBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (lastBlobUrl && lastOutName) triggerDownload(lastBlobUrl, lastOutName);
        });

        resetBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            fileInput.value = '';
            successState.classList.add('hidden');
            successState.classList.remove('flex');
            idleState.classList.remove('hidden');
        });

        function triggerDownload(url, filename) {
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
        }

        async function handleFile(file) {
            idleState.classList.add('hidden');
            successState.classList.add('hidden');
            successState.classList.remove('flex');
            loadingState.classList.remove('hidden');
            loadingState.classList.add('flex');

            progressBar.style.width = '30%';
            statusText.innerText = `Uploading ${file.name}...`;

            const formData = new FormData();
            formData.append('file', file);
            try {
                formData.append('tz', Intl.DateTimeFormat().resolvedOptions().timeZone || 'Asia/Kolkata');
            } catch (e) {
                formData.append('tz', 'Asia/Kolkata');
            }

            const progressTimer = setTimeout(() => {
                progressBar.style.width = '75%';
                statusText.innerText = 'Transliterating via Sanka + Sanchaya Engine...';
            }, 500);

            try {
                const response = await fetch('/convert', { method: 'POST', body: formData });
                clearTimeout(progressTimer);

                if (!response.ok) throw new Error('Conversion failed');

                progressBar.style.width = '100%';
                const blob = await response.blob();
                const lower = file.name.toLowerCase();
                const ext = (lower.endsWith('.doc') || lower.endsWith('.docx')) ? '.docx' : '.xlsx';
                const dotIndex = file.name.lastIndexOf('.');
                const baseName = dotIndex !== -1 ? file.name.substring(0, dotIndex) : file.name;
                const outName = `kannada_unicode_${baseName}${ext}`;

                if (lastBlobUrl) window.URL.revokeObjectURL(lastBlobUrl);
                lastBlobUrl = window.URL.createObjectURL(blob);
                lastOutName = outName;

                triggerDownload(lastBlobUrl, lastOutName);

                loadingState.classList.add('hidden');
                loadingState.classList.remove('flex');
                successState.classList.remove('hidden');
                successState.classList.add('flex');
                document.getElementById('convertedFileName').innerText = outName;
            } catch (err) {
                clearTimeout(progressTimer);
                alert('⚠️ Conversion error or unsupported file. Please try again.');
                loadingState.classList.add('hidden');
                loadingState.classList.remove('flex');
                idleState.classList.remove('hidden');
            }
        }
    </script>
</body>
</html>
    """
