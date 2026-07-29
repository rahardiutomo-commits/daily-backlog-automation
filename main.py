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

# Masukkan ID Spreadsheet & Folder Drive kamu
SPREADSHEET_ID = "1oI-f_KPFq..."  # <-- SESUAIKAN DENGAN ID SPREADSHEET KAMU
DRIVE_FOLDER_ID = "1zwQoKtOK..."  # <-- SESUAIKAN DENGAN ID FOLDER DRIVE KAMU

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
        # Tambahkan label khusus Gmail
        mail.store(email_id, '+X-GM-LABELS', 'PROCESSED_BACKLOG')
        # Tandai terbaca (\Seen)
        mail.store(email_id, '+FLAGS', '\\Seen')
        mail.logout()
        print("✅ Email berhasil diberi label & ditandai terbaca!")
    except Exception as e:
        print(f"⚠️ Gagal memberi label email: {str(e)}")

# ==========================================
# 3. PROSES DATA & PENULISAN SHEETS (15 KOLOM A-O)
# ==========================================
def process_and_update_sheets(csv_bytes):
    print("📊 Menganalisis data & memperbarui Google Sheets...")
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet("BACKLOG_HISTORY") # <-- Sesuaikan nama Tab kamu

    # Buka CSV dengan pandas
    df = pd.read_csv(io.BytesIO(csv_bytes))
    df.columns = [c.strip() for c.columns]

    today_date = pd.Timestamp.now().strftime('%Y-%m-%d')

    # Pemrosesan LOB / Data (Sesuaikan jika kamu punya pemisahan per LOB)
    # Contoh struktur susunan 15 Kolom yang PRESISI A - O:
    
    # Keterangan Kolom:
    # A: Tanggal backlog, B: Total Backlog, C: LOB, D: Top mobile sub topic, E: Case
    # F: 0-3 Days, G: 3-7, H: 7-14, I: >14
    # J: AGING >30, K: AGING >60, L: Analisa from yesterday, M: TREND, N: CHANGE, O: INSIGHT

    total_backlog = len(df)
    
    # *Contoh penyusunan array baris data 15 elemen*
    row_data = [
        today_date,                     # A: Tanggal
        total_backlog,                  # B: Total Backlog
        "Fraud / Non-Fraud / Channel",  # C: LOB
        "- TOP_TOPIC_1: 100",           # D: Top mobile sub topic
        "- CASE_1: 50",                 # E: Case
        100,                            # F: 0-3 Days
        50,                             # G: 3-7
        20,                             # H: 7-14
        10,                             # I: >14
        "Detail / List Tiket >30 Hari", # J: AGING >30
        "Detail / List Tiket >60 Hari", # K: AGING >60
        "Analisa perubahan backlog...", # L: Analisa from yesterday
        "UP 📈",                        # M: TREND
        "+152",                         # N: CHANGE
        "Insight fokus area..."         # O: INSIGHT
    ]

    worksheet.append_row(row_data, value_input_option="USER_ENTERED")
    print("✅ Berhasil menulis baris baru sesuai urutan Kolom A sampai O!")

# ==========================================
# 4. EXECUTION MAIN
# ==========================================
if __name__ == "__main__":
    try:
        csv_bytes, file_name, mail, email_id = download_latest_csv_from_gmail()
        
        if csv_bytes is None:
            print("🚀 Tidak ada email baru. Workflow selesai.")
            exit(0)

        process_and_update_sheets(csv_bytes)
        mark_email_as_processed(mail, email_id)
        
        print("🎉 AUTOMATION BERHASIL SELESAI!")
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        exit(1)
