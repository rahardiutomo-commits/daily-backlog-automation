import os
import io
import json
import imaplib
import email
from email.header import decode_header
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ==========================================
# 1. KREDENSIAL DARI ENVIRONMENT
# ==========================================
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
GCP_SA_KEY = os.environ.get("GCP_SA_KEY")

if not all([GMAIL_USER, GMAIL_APP_PASSWORD, GCP_SA_KEY]):
    raise ValueError("⚠️ Secret GitHub belum lengkap! Cek GMAIL_USER, GMAIL_APP_PASSWORD, dan GCP_SA_KEY.")

# Konfigurasi Google Credentials
sa_info = json.loads(GCP_SA_KEY)
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
credentials = Credentials.from_service_account_info(sa_info, scopes=scopes)
gc = gspread.authorize(credentials)
drive_service = build('drive', 'v3', credentials=credentials)

# ID ASLI BERDASARKAN LINK KAMU:
SPREADSHEET_ID = "1oI-f_KPFqTwe8Q0M3zva1f2QbbsBeegDi-7Yly-W1Cs"
DRIVE_FOLDER_ID = "1zwQoKtOKf0houdAFIgC1MADFU8oygvt_"

# ==========================================
# 2. AMBIL EMAIL & LABEL GMAIL AUTOMATION
# ==========================================
def download_latest_csv_from_gmail():
    print("📧 Memeriksa email masuk di Gmail...")
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    # Cari email yang BELUM memiliki label 'PROCESSED_BACKLOG'
    status, messages = mail.search(None, 'NOT', 'X-GM-LABELS', 'PROCESSED_BACKLOG')
    mail_ids = messages[0].split()

    if not mail_ids:
        print("ℹ️ Tidak ada email baru yang perlu diproses hari ini.")
        mail.logout()
        return None, None, None, None

    # Ambil email terbaru yang belum diproses
    latest_email_id = mail_ids[-1]
    status, msg_data = mail.fetch(latest_email_id, "(RFC822)")

    csv_content = None
    file_name = "backlog_data.csv"

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                
                filename = part.get_filename()
                if filename and (filename.endswith('.csv') or filename.endswith('.xls')):
                    csv_content = part.get_payload(decode=True)
                    file_name = filename
                    break

    if not csv_content:
        mail.logout()
        raise Exception("❌ Email ditemukan tetapi tidak memiliki lampiran CSV!")

    print(f"✅ Berhasil mengunduh lampiran email: {file_name}")
    return csv_content, file_name, mail, latest_email_id


def mark_email_as_processed(mail, email_id):
    """Menambahkan label PROCESSED_BACKLOG ke Gmail via IMAP X-GM-LABELS"""
    try:
        print("🏷️ Menandai email dengan label 'PROCESSED_BACKLOG'...")
        mail.store(email_id, '+X-GM-LABELS', 'PROCESSED_BACKLOG')
        mail.store(email_id, '+FLAGS', '\\Seen')
        mail.logout()
        print("✅ Email berhasil diberi label & ditandai terbaca!")
    except Exception as e:
        print(f"⚠️ Gagal memberi label email: {str(e)}")

# ==========================================
# 3. PROSES DATA & PENULISAN SHEETS (15 KOLOM A-O)
# ==========================================
def process_and_update_sheets(csv_bytes, file_name):
    print("📊 Menganalisis data & memperbarui Google Sheets...")
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet("BACKLOG_HISTORY") # Tab BACKLOG_HISTORY

    # Buka CSV dengan pandas
    df = pd.read_csv(io.BytesIO(csv_bytes))
    df.columns = [c.strip() for c.columns]

    today_date = pd.Timestamp.now().strftime('%Y-%m-%d')
    total_backlog = len(df)

    # Membaca baris terakhir kemarin untuk kalkulasi Trend & Change
    all_rows = worksheet.get_all_values()
    yesterday_total = total_backlog
    if len(all_rows) > 1:
        last_val = all_rows[-1][1] # Kolom B (Total Backlog)
        if last_val.isdigit():
            yesterday_total = int(last_val)

    diff = total_backlog - yesterday_total
    if diff > 0:
        trend = "UP 📈"
        change_str = f"+{diff:,}"
    elif diff < 0:
        trend = "DOWN 📉"
        change_str = f"{diff:,}"
    else:
        trend = "STABLE ➖"
        change_str = "0"

    pct_change = ((diff / yesterday_total) * 100) if yesterday_total > 0 else 0

    analisa_yesterday = (
        f"1. Backlog {('naik' if diff >= 0 else 'turun')} {abs(pct_change):.1f}% ({change_str} tiket) dibanding laporan sebelumnya.\n"
        f"2. Eskalasi penanganan tiket aging perlu dijaga ketat."
    )
    insight = "💡 Focus Area: Percepat penyelesaian tiket berumur >14 hari."

    # Susun 15 Kolom Presisi (A sampai O)
    row_data = [
        today_date,          # A: Tanggal backlog
        total_backlog,       # B: Total Backlog
        "General",           # C: LOB
        "- TOP_ISSUE: 50",   # D: Top mobile sub topic
        "- CASE_1: 20",      # E: Case
        0,                   # F: Aging 0-3 Days
        0,                   # G: 3-7
        0,                   # H: 7-14
        0,                   # I: >14
        "List Tiket >30",    # J: AGING >30
        "List Tiket >60",    # K: AGING >60
        analisa_yesterday,   # L: Analisa from yesterday
        trend,               # M: TREND
        change_str,          # N: CHANGE
        insight              # O: INSIGHT
    ]

    worksheet.append_row(row_data, value_input_option="USER_ENTERED")
    print("✅ Berhasil menulis ke tab BACKLOG_HISTORY!")

# ==========================================
# 4. UPLOAD CSV KE DRIVE
# ==========================================
def upload_to_drive(csv_bytes, file_name):
    print("☁️ Mengunggah backup CSV ke Google Drive...")
    file_metadata = {
        'name': f"{pd.Timestamp.now().strftime('%Y%m%d')}_{file_name}",
        'parents': [DRIVE_FOLDER_ID]
    }
    media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype='text/csv', resumable=True)
    drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print("✅ Backup CSV tersimpan di Drive!")

# ==========================================
# 5. EXECUTION MAIN
# ==========================================
if __name__ == "__main__":
    try:
        csv_bytes, file_name, mail, email_id = download_latest_csv_from_gmail()
        
        if csv_bytes is None:
            print("🚀 Tidak ada email baru. Automation selesai.")
            exit(0)

        process_and_update_sheets(csv_bytes, file_name)
        upload_to_drive(csv_bytes, file_name)
        mark_email_as_processed(mail, email_id)
        
        print("🎉 AUTOMATION SELESAI DENGAN SUKSES!")
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        exit(1)
