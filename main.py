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

SPREADSHEET_ID = "1oI-f_KPFqTwe8Q0M3zva1f2QbbsBeegDi-7Yly-W1Cs"
DRIVE_FOLDER_ID = "1zwQoKtOKf0houdAFIgC1MADFU8oygvt_"

# ==========================================
# 2. EMAIL & GMAIL LABEL
# ==========================================
def download_latest_csv_from_gmail():
    print("📧 Memeriksa email masuk di Gmail...")
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    # Cari email yang BELUM berlabel PROCESSED_BACKLOG
    status, messages = mail.search(None, 'NOT', 'X-GM-LABELS', 'PROCESSED_BACKLOG')
    mail_ids = messages[0].split()

    if not mail_ids:
        print("ℹ️ Tidak ada email baru yang perlu diproses hari ini.")
        mail.logout()
        return None, None, None, None

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
        raise Exception("❌ Email ditemukan tetapi tidak ada lampiran CSV!")

    print(f"✅ Berhasil mengambil lampiran email: {file_name}")
    return csv_content, file_name, mail, latest_email_id

def mark_email_as_processed(mail, email_id):
    try:
        print("🏷️ Menandai email dengan label 'PROCESSED_BACKLOG'...")
        mail.store(email_id, '+X-GM-LABELS', 'PROCESSED_BACKLOG')
        mail.store(email_id, '+FLAGS', '\\Seen')
        mail.logout()
        print("✅ Email berhasil diberi label & ditandai terbaca!")
    except Exception as e:
        print(f"⚠️ Gagal memberi label email: {str(e)}")

# ==========================================
# 3. PROSES DATA PER LOB & UPDATE SHEETS
# ==========================================
def process_and_update_sheets(csv_bytes, file_name):
    print("📊 Menganalisis data & memperbarui Google Sheets...")
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet("BACKLOG_HISTORY")

    df = pd.read_csv(io.BytesIO(csv_bytes))
    df.columns = [c.strip() for c.columns]

    today_date = pd.Timestamp.now().strftime('%Y-%m-%d')

    # Pemetaan LOB
    lobs = ['Fraud', 'Non Fraud', 'Channel']
    
    for lob in lobs:
        # Filter dataframe per LOB jika ada kolom LOB/Category di CSV
        if 'LOB' in df.columns:
            sub_df = df[df['LOB'].str.contains(lob, case=False, na=False)]
        elif 'Category' in df.columns:
            sub_df = df[df['Category'].str.contains(lob, case=False, na=False)]
        else:
            sub_df = df # Fallback jika tidak ada kolom LOB

        total_backlog = len(sub_df)
        if total_backlog == 0 and not df.empty:
            continue

        # Top Mobile Sub Topic & Case
        top_topic_str = ""
        if 'Sub Topic' in sub_df.columns and not sub_df.empty:
            top_topics = sub_df['Sub Topic'].value_counts().head(5)
            top_topic_str = "\n".join([f"- {k}: {v}" for k, v in top_topics.items()])

        top_case_str = ""
        if 'Case' in sub_df.columns and not sub_df.empty:
            top_cases = sub_df['Case'].value_counts().head(5)
            top_case_str = "\n".join([f"- {k}: {v}" for k, v in top_cases.items()])

        # Hitung Aging
        aging_0_3 = len(sub_df[sub_df['Aging'] <= 3]) if 'Aging' in sub_df.columns else 0
        aging_3_7 = len(sub_df[(sub_df['Aging'] > 3) & (sub_df['Aging'] <= 7)]) if 'Aging' in sub_df.columns else 0
        aging_7_14 = len(sub_df[(sub_df['Aging'] > 7) & (sub_df['Aging'] <= 14)]) if 'Aging' in sub_df.columns else 0
        aging_14_plus = len(sub_df[sub_df['Aging'] > 14]) if 'Aging' in sub_df.columns else 0

        # Ringkasan AGING >30 dan >60 (Singkat agar tidak merusak kolom J & K)
        cnt_30 = len(sub_df[sub_df['Aging'] > 30]) if 'Aging' in sub_df.columns else 0
        cnt_60 = len(sub_df[sub_df['Aging'] > 60]) if 'Aging' in sub_df.columns else 0

        text_aging_30 = f"Ticket >30 Hari: {cnt_30}"
        text_aging_60 = f"Ticket >60 Hari: {cnt_60}"

        # Analisa & Insight untuk LOB ini
        analisa_text = (
            f"1. Total backlog {lob}: {total_backlog} tiket.\n"
            f"2. Tiket berumur >14 hari sebanyak {aging_14_plus} tiket.\n"
            f"3. Kontributor terbesar didominasi sub-topic terkait."
        )
        trend_status = "STABLE ➖"
        change_val = "0"
        insight_text = f"💡 Focus LOB {lob}: Percepat SLA tiket aging."

        # SUSUN 15 KOLOM PRESISI (A - O)
        row_data = [
            today_date,      # A: Tanggal backlog
            total_backlog,   # B: Total Backlog
            lob,             # C: LOB (Fraud / Non Fraud / Channel)
            top_topic_str,   # D: Top mobile sub topic
            top_case_str,    # E: Case
            aging_0_3,       # F: Aging ticket 0-3 Days
            aging_3_7,       # G: 3-7
            aging_7_14,      # H: 7-14
            aging_14_plus,   # I: >14
            text_aging_30,   # J: AGING >30
            text_aging_60,   # K: AGING >60
            analisa_text,    # L: Analisa from yesterday
            trend_status,    # M: TREND
            change_val,      # N: CHANGE
            insight_text     # O: INSIGHT
        ]

        worksheet.append_row(row_data, value_input_option="USER_ENTERED")

    print("✅ Berhasil menulis baris per LOB pas di Kolom A sampai O!")

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
            print("🚀 Tidak ada email baru. Workflow selesai.")
            exit(0)

        process_and_update_sheets(csv_bytes, file_name)
        upload_to_drive(csv_bytes, file_name)
        mark_email_as_processed(mail, email_id)
        
        print("🎉 AUTOMATION SELESAI DENGAN SUKSES!")
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        exit(1)
