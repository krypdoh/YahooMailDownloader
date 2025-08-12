# YahooMailDownloader
Connect to a Yahoo Mail account via IMAP using SSL and an app password. Downloads emails and attachments and deletes from server.
Yahoo Mail Email Downloader and Archiver

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

USE AT YOUR OWN RISK!! PLEASE TEST IT ON SAMPLE FOLDER OR ACCOUNT BEFORE PROCEEDING! 

<img width="1935" height="715" alt="image" src="https://github.com/user-attachments/assets/fd8f93d0-f890-4f92-88ea-2e887bd1720d" />

<img width="2526" height="715" alt="image" src="https://github.com/user-attachments/assets/4e289e56-7f7a-4e1a-9e4c-18aceb74b86d" />

