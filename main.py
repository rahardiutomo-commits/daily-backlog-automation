import os
import io
import json
import re
import imaplib
import email
from email.header import decode_header
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ==========================================
# 1. AMBIL KREDENSIAL DARI ENVIRONMENT
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

# ID Google Sheet & Drive Folder kamu
SPREADSHEET_ID = "1oI-f_KPFq..."  # <-- Sesuaikan ID Sheet kamu
DRIVE_FOLDER_ID = "1zwQoKtOK..."  # <-- Sesuaikan ID Folder Drive kamu

# ==========================================
# 2. AMBIL FILE CSV DARI GMAIL
# ==========================================
def download_latest_csv_from_gmail():
    print("📧 Merekam email terbaru dari Gmail...")
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    # Cari email dari Kapture / Backlog Report
    status, messages = mail.search(None, 'ALL')
    mail_ids = messages[0].split()
    
    if not mail_ids:
        raise Exception("Tidak ada email ditemukan di Inbox.")

    # Ambil email paling terakhir
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

    mail.logout()
    if not csv_content:
        raise Exception("❌ Lampiran CSV tidak ditemukan pada email terbaru!")
    
    print(f"✅ Berhasil mengunduh lampiran: {file_name}")
    return csv_content, file_name

# ==========================================
# 3. LOGIKA ANALISIS SELEPAS MANAGER OPS
# ==========================================
def process_and_analyze_backlog(csv_bytes):
    print("📊 Memproses & menganalisis data backlog...")
    df = pd.read_csv(io.BytesIO(csv_bytes))

    # --- A. Pembersihan Data Dasar ---
    # *Pastikan nama kolom di bawah sesuai dengan header CSV Kapture kamu*
    df.columns = [c.strip() for c.columns] 
    
    total_backlog = len(df)
    
    # Hitung status tiket
    open_tickets = len(df[df['Status'].str.contains('Open', case=False, na=False)]) if 'Status' in df.columns else 0
    pending_vendor = len(df[df['Status'].str.contains('Vendor', case=False, na=False)]) if 'Status' in df.columns else 0
    pending_user = len(df[df['Status'].str.contains('User', case=False, na=False)]) if 'Status' in df.columns else 0

    # --- B. Segmentasi Aging Tiket ---
    # Asumsi ada kolom 'Aging' (dalam hari)
    if 'Aging' in df.columns:
        aging_0_7 = len(df[df['Aging'] <= 7])
        aging_7_14 = len(df[(df['Aging'] > 7) & (df['Aging'] <= 14)])
        aging_14_30 = len(df[(df['Aging'] > 14) & (df['Aging'] <= 30)])
        
        df_30_60 = df[(df['Aging'] > 30) & (df['Aging'] <= 60)]
        df_over_60 = df[df['Aging'] > 60]
    else:
        aging_0_7 = aging_7_14 = aging_14_30 = 0
        df_30_60 = pd.DataFrame()
        df_over_60 = pd.DataFrame()

    # --- C. Format Teks Detail Tiket (Aging >30 & >60) ---
    def format_ticket_list(sub_df, title, max_items=8):
        count = len(sub_df)
        if count == 0:
            return f"⚠️ {title} : 0\n\nNihil / Clean"
        
        res = f"⚠️ {title} : {count}\n\n📌 Persisten / Kritis ({count})\n"
        for _, row in sub_df.head(max_items).iterrows():
            t_id = str(row.get('Ticket ID', 'ID_Unknown'))
            days = str(row.get('Aging', '-'))
            sub_topic = str(row.get('Sub Topic', 'General'))
            res += f"• {t_id} | {days} Hari | {sub_topic}\n"
        return res

    text_aging_30_60 = format_ticket_list(df_30_60, "Ticket >30 Hari")
    text_aging_over_60 = format_ticket_list(df_over_60, "Ticket >60 Hari")

    # --- D. Top Contributor Sub-Topic ---
    top_issue = "N/A"
    top_issue_count = 0
    if 'Sub Topic' in df.columns and not df.empty:
        top_series = df['Sub Topic'].value_counts()
        if not top_series.empty:
            top_issue = top_series.index[0]
            top_issue_count = top_series.iloc[0]

    return {
        "total": total_backlog,
        "open": open_tickets,
        "pending_vendor": pending_vendor,
        "pending_user": pending_user,
        "aging_0_7": aging_0_7,
        "aging_7_14": aging_7_14,
        "aging_14_30": aging_14_30,
        "text_aging_30": text_aging_30_60,
        "text_aging_60": text_aging_over_60,
        "top_issue": top_issue,
        "top_issue_count": top_issue_count
    }

# ==========================================
# 4. UPDATE GOOGLE SHEETS & KALKULASI OPS
# ==========================================
def update_google_sheets(metrics):
    print("📑 Mengisi ke Google Sheets...")
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet("Raw Data") # <-- Nama Tab kamu

    # 1. Ambil data baris terakhir (kemarin) untuk pembanding
    all_rows = worksheet.get_all_values()
    last_row = all_rows[-1] if len(all_rows) > 1 else None

    yesterday_total = int(last_row[1]) if (last_row and last_row[1].isdigit()) else metrics["total"]

    # 2. Hitung Trend, Selisih (Change), dan Persentase Perubahan
    diff = metrics["total"] - yesterday_total
    
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

    # 3. Buat Narasi Laporan (Analisa From Yesterday & Insight)
    today_date = pd.Timestamp.now().strftime('%Y-%m-%d')
    
    analisa_yesterday = (
        f"1. Backlog {('naik' if diff >= 0 else 'turun')} {abs(pct_change):.1f}% ({change_str} tiket) dibanding laporan sebelumnya.\n"
        f"2. Kenaikan/Keterlambatan tertinggi didominasi oleh sub-topic '{metrics['top_issue']}' ({metrics['top_issue_count']} tiket).\n"
        f"3. Kontributor terbanyak perlu eskalasi harian agar tidak menumpuk ke Aging >30."
    )

    insight = (
        f"💡 Focus Area: Penanganan pada kategori '{metrics['top_issue']}' perlu dipercepat. "
        f"Pastikan tiket pending vendor diawasi H+1 agar SLA tidak breaching."
    )

    # 4. Susun Baris Data Sesuai Kolom A sampai O
    new_row = [
        today_date,                   # A: Date
        metrics["total"],             # B: Total Backlog
        metrics["pending_vendor"],    # C: Pending Vendor
        metrics["pending_user"],      # D: Pending User
        metrics["open"],              # E: Open
        metrics["aging_0_7"],         # F: 0-7 Hari
        metrics["aging_7_14"],        # G: 7-14 Hari
        metrics["aging_14_30"],       # H: >14 Hari
        metrics["text_aging_30"],     # I: AGING >30 (Detail Tiket)
        "",                           # J: (KOSONGKAN / PADDING)
        metrics["text_aging_60"],     # K: AGING >60 (Detail Tiket)
        analisa_yesterday,            # L: Analisa from yesterday
        trend,                        # M: TREND
        change_str,                   # N: CHANGE
        insight                       # O: INSIGHT
    ]

    # Append ke Google Sheet
    worksheet.append_row(new_row, value_input_option="USER_ENTERED")
    print("✅ Berhasil update Google Sheet lengkap A-O!")

# ==========================================
# 5. UPLOAD FILE CSV KE GOOGLE DRIVE
# ==========================================
def upload_to_drive(csv_bytes, file_name):
    print("☁️ Mengunggah backup CSV ke Google Drive...")
    file_metadata = {
        'name': f"{pd.Timestamp.now().strftime('%Y%m%d')}_{file_name}",
        'parents': [DRIVE_FOLDER_ID]
    }
    media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype='text/csv', resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"✅ Backup tersimpan di Drive dengan File ID: {file.get('id')}")

# ==========================================
# MAIN EXECUTION
# ==========================================
if __name__ == "__main__":
    try:
        csv_bytes, file_name = download_latest_csv_from_gmail()
        metrics = process_and_analyze_backlog(csv_bytes)
        update_google_sheets(metrics)
        upload_to_drive(csv_bytes, file_name)
        print("🎉 AUTOMATION SELESAI DENGAN SUKSES!")
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        exit(1)
