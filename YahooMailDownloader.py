"""
===================================================================================
Yahoo Mail Email Downloader
-----------------------------------------------------------------------------------
This script connects to a Yahoo Mail account via IMAP using SSL and an app password.

Functionality:
-------------
- Searches and downloads all emails from a specified year.
- Saves each email as a .eml file in a year-based folder under the SAVE_DIR.
- Extracts and saves attachments (if present) in the same folder as the email.
- Deletes the email from the server after successful download.
- Repeats the download process until no emails remain for the selected year.
- Supports optional multithreaded downloading (configurable).
- Reconnects to the server automatically on connection failures or SSL errors.

Requirements:
-------------
- Python 3.6+
- Yahoo app password (2FA must be enabled in Yahoo)
- Stable internet connection

Configuration:
--------------
- Set USERNAME and APP_PASSWORD to your Yahoo credentials.
- Set SAVE_DIR to the directory where emails should be archived.

Usage:
------
1. Run the script.
2. Enter the year you want to archive (e.g., 2024).
3. The script will:
   - Find all emails from that year,
   - Save them and their attachments,
   - Delete them from Yahoo Mail,
   - Repeat until no emails are left for that year.

IMPORTANT:
----------
- This script permanently deletes emails from your Yahoo account after saving.
- Use with caution and verify your backups.

Author: Krypdoh
Last updated: 2025.07.31

===================================================================================
"""

import imaplib
import email
import os
from email.utils import parsedate_tz
from email.header import decode_header
import time
import ssl
import datetime
import re

# === Configuration ===
USERNAME = 'email@yahoo.com'
APP_PASSWORD = 'yahoo_app_password'  # Use an app-specific password
SAVE_DIR = r'D:\YahooEmails'

os.makedirs(SAVE_DIR, exist_ok=True)

def clean(text):
    """Clean text for safe filenames but preserve file extensions."""
    # Keep alphanumerics, underscore, hyphen, dot
    return re.sub(r'[^\w\.-]', '_', text)

def decode_mime_header(header_value):
    """Decode MIME encoded headers to readable string."""
    if not header_value:
        return ""
    decoded_parts = decode_header(header_value)
    decoded_str = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            decoded_str += part.decode(encoding or 'utf-8', errors='ignore')
        else:
            decoded_str += part
    return decoded_str

def reconnect_mail():
    """Reconnect to Yahoo IMAP with retry."""
    context = ssl.create_default_context()
    context.set_ciphers('HIGH:!aNULL:!eNULL')

    for attempt in range(3):
        try:
            mail = imaplib.IMAP4_SSL("export.imap.mail.yahoo.com", ssl_context=context, timeout=30)
            mail.login(USERNAME, APP_PASSWORD)
            mail.select("inbox")
            return mail
        except Exception as e:
            print(f"Reconnect attempt {attempt + 1}/3 failed: {e}")
            time.sleep(3)

    raise Exception("Unable to reconnect after 3 attempts.")

def save_email(email_message, year_folder):
    """Save email and attachments to disk."""
    folder_path = os.path.join(SAVE_DIR, year_folder)
    os.makedirs(folder_path, exist_ok=True)

    # Create base filename
    msg_id = email_message.get('Message-ID')
    if msg_id:
        base_name = clean(msg_id)
    else:
        date_str = email_message.get('Date', '').replace(' ', '_').replace(':', '-')
        subject = decode_mime_header(email_message.get('Subject', 'No_Subject'))
        base_name = clean(f"{date_str}_{subject}")

    # Save the raw email .eml
    eml_path = os.path.join(folder_path, base_name + ".eml")
    with open(eml_path, "wb") as f:
        f.write(email_message.as_bytes())

    # Walk email parts to find attachments
    for part in email_message.walk():
        content_disposition = part.get("Content-Disposition", "")
        content_disposition = decode_mime_header(content_disposition)
        
        filename = part.get_filename()
        filename = decode_mime_header(filename)

        if filename or ("attachment" in content_disposition.lower()):
            if filename:
                safe_filename = clean(filename)
            else:
                safe_filename = "attachment.bin"
            attachment_path = os.path.join(folder_path, safe_filename)
            
            try:
                payload = part.get_payload(decode=True)
                if payload:
                    with open(attachment_path, "wb") as f:
                        f.write(payload)
                    print(f"📎 Attachment saved: {attachment_path}")
            except Exception as e:
                print(f"Failed to save attachment {safe_filename}: {e}")

def process_email(mail, email_id, retries=2):
    """Fetch, save, and delete an individual email with retry."""
    try:
        result, msg_data = mail.fetch(email_id, "(RFC822)")
        if result != "OK":
            return f"Failed to fetch email {email_id.decode()}"

        raw_email = msg_data[0][1]
        email_message = email.message_from_bytes(raw_email)

        date_tuple = parsedate_tz(email_message.get('Date'))
        year = str(date_tuple[0]) if date_tuple else "Unknown_Year"

        save_email(email_message, year)

        # Delete email from server
        mail.store(email_id, '+FLAGS', r'(\Deleted)')
        mail.expunge()

        return f"✅ Email {email_id.decode()} saved to {year} and deleted"

    except (imaplib.IMAP4.abort, ssl.SSLError, OSError) as e:
        if retries > 0:
            print(f"⚠️ Connection error on email {email_id.decode()}, retrying... ({retries} left)")
            time.sleep(3)
            new_mail = reconnect_mail()
            return process_email(new_mail, email_id, retries - 1)
        return f"❌ Failed email {email_id.decode()} after retries: {e}"

    except Exception as e:
        return f"❌ Error processing email {email_id.decode()}: {e}"

def format_imap_date(date_str):
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    return dt.strftime("%d-%b-%Y")

def download_yahoo_emails(year_filter):
    """Download and delete all emails for a given year. Return number of emails processed."""
    mail = reconnect_mail()

    since_date = format_imap_date(f"{year_filter}-01-01")
    before_date = format_imap_date(f"{int(year_filter) + 1}-01-01")

    result, data = mail.search(None, 'SINCE', since_date, 'BEFORE', before_date)

    if result != "OK":
        print("No messages found!")
        return 0

    email_ids = data[0].split()
    count = len(email_ids)
    print(f"📬 Found {count} messages. Processing...")

    # Sequential processing
    for email_id in email_ids:
        print(process_email(mail, email_id))

    return count

if __name__ == "__main__":
    year_filter = input("Enter the year to filter emails by (e.g. 2024): ").strip()
    if not year_filter.isdigit() or len(year_filter) != 4:
        print("❌ Please enter a valid 4-digit year.")
    else:
        while True:
            count = download_yahoo_emails(year_filter)
            if count == 0:
                print(f"✅ All emails from {year_filter} have been processed.")
                break
            else:
                print(f"🔁 {count} emails processed. Checking again in 5 seconds...")
                time.sleep(5)
