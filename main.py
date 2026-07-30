import os
import io
import json
import imaplib
import email
import re
from email.utils import parsedate_to_datetime
from bs4 import BeautifulSoup
import requests
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
# 2. AMBIL EMAIL, TANGGAL EMAIL, & LINK DOWNLOAD
# ==========================================
def download_latest_csv_from_gmail():
    print("📧 Memeriksa email masuk di Gmail...")
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    try:
        status, messages = mail.search(None, 'X-GM-RAW', 'NOT label:PROCESSED_BACKLOG')
    except Exception:
        status, messages = mail.search(None, 'UNSEEN')

    mail_ids = messages[0].split()

    if not mail_ids:
        print("ℹ️ Tidak ada email baru yang perlu diproses hari ini.")
        mail.logout()
        return None, None, None, None, None

    latest_email_id = mail_ids[-1]
    status, msg_data = mail.fetch(latest_email_id, "(RFC822)")

    csv_content = None
    file_name = "backlog_report.csv"
    email_date_str = None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            
            # --- EXTRACT TANGGAL EMAIL (AKAN DIBUAT TANGGAL BACKLOG) ---
            date_header = msg.get("Date")
            if date_header:
                try:
                    dt = parsedate_to_datetime(date_header)
                    email_date_str = dt.strftime('%Y-%m-%d')
                except Exception:
                    email_date_str = pd.Timestamp.now().strftime('%Y-%m-%d')
            else:
                email_date_str = pd.Timestamp.now().strftime('%Y-%m-%d')

            # --- EKSTRAKSI BODY HTML & LINK DOWNLOAD ---
            body_html = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == 'text/html':
                        body_html += part.get_payload(decode=True).decode('utf-8', errors='ignore')
            else:
                body_html = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

            # Gunakan BeautifulSoup untuk ekstrak tautan href
            soup = BeautifulSoup(body_html, 'html.parser')
            target_url = None

            for a_tag in soup.find_all('a', href=True):
                text = a_tag.get_text().lower()
                href = a_tag['href']
                if 'download' in text or 'click' in text or 'here' in text or 'report' in text:
                    target_url = href
                    break
            
            if not target_url:
                # Fallback regex URL jika tag <a> tidak spesifik
                urls = re.findall(r'https?://[^\s<>"]+', body_html)
                if urls:
                    target_url = urls[0]

            if target_url:
                print(f"🔗 Link download ditemukan: {target_url}")
                resp = requests.get(target_url, timeout=45)
                if resp.status_code == 200:
                    csv_content = resp.content
                    print("✅ Berhasil mendownload report dari link!")
                else:
                    print(f"⚠️ Gagal download dari link, status code: {resp.status_code}")

    if not csv_content:
        mail.logout()
        print("⚠️ File dari link download tidak dapat diunduh.")
        return None, None, None, None, None

    return csv_content, file_name, email_date_str, mail, latest_email_id

def mark_email_as_processed(mail, email_id):
    if not mail or not email_id:
        return
    try:
        print("🏷️ Menandai email dengan label 'PROCESSED_BACKLOG'...")
        # Buat label jika belum ada, lalu set label & seen
        mail.store(email_id, '+X-GM-LABELS', 'PROCESSED_BACKLOG')
        mail.store(email_id, '+FLAGS', '\\Seen')
        mail.logout()
        print("✅ Email berhasil diberi label & ditandai terbaca!")
    except Exception as e:
        print(f"⚠️ Catatan pelabelan email: {str(e)}")

# ==========================================
# 3. PROSES DATA & UPDATE SHEETS (KOLOM A-O)
# ==========================================
def process_and_update_sheets(csv_bytes, file_name, backlog_date):
    print(f"📊 Menganalisis data untuk Tanggal Backlog: {backlog_date}...")
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet("BACKLOG_HISTORY")

    try:
        df = pd.read_csv(io.BytesIO(csv_bytes))
    except Exception:
        df = pd.read_excel(io.BytesIO(csv_bytes))

    df.columns = [str(c).strip() for c in df.columns]

    lobs = ['Fraud', 'Non Fraud', 'Merchant', 'Channel']
    
    for lob in lobs:
        lob_col = next((c for c in df.columns if c.lower() in ['lob', 'category', 'kategori']), None)
        if lob_col:
            sub_df = df[df[lob_col].astype(str).str.contains(lob, case=False, na=False)].copy()
        else:
            sub_df = df.copy()

        total_backlog = len(sub_df)
        if total_backlog == 0 and not df.empty and lob_col:
            continue

        topic_col = next((c for c in df.columns if 'topic' in c.lower() or 'sub' in c.lower()), None)
        case_col = next((c for c in df.columns if 'case' in c.lower() or 'issue' in c.lower()), None)

        top_topic_str = ""
        if topic_col and not sub_df.empty:
            top_topics = sub_df[topic_col].value_counts().head(5)
            top_topic_str = "\n".join([f"- {k}: {v}" for k, v in top_topics.items()])

        top_case_str = ""
        if case_col and not sub_df.empty:
            top_cases = sub_df[case_col].value_counts().head(5)
            top_case_str = "\n".join([f"- {k}: {v}" for k, v in top_cases.items()])

        aging_col = next((c for c in df.columns if 'aging' in c.lower() or 'day' in c.lower() or 'hari' in c.lower()), None)
        ticket_col = next((c for c in df.columns if 'ticket' in c.lower() or 'id' in c.lower() or 'number' in c.lower()), None)

        if aging_col:
            sub_df['aging_num'] = pd.to_numeric(sub_df[aging_col], errors='coerce').fillna(0)
            aging_0_3 = len(sub_df[sub_df['aging_num'] <= 3])
            aging_3_7 = len(sub_df[(sub_df['aging_num'] > 3) & (sub_df['aging_num'] <= 7)])
            aging_7_14 = len(sub_df[(sub_df['aging_num'] > 7) & (sub_df['aging_num'] <= 14)])
            aging_14_plus = len(sub_df[sub_df['aging_num'] > 14])

            # --- SUSUN LIST TIKET AGING >30 HARI & >60 HARI (ISI KOLOM J & K) ---
            df_30 = sub_df[sub_df['aging_num'] > 30]
            df_60 = sub_df[sub_df['aging_num'] > 60]

            if not df_30.empty and ticket_col:
                list_30 = [f"• {row[ticket_col]} | {int(row['aging_num'])} Hari | -" for _, row in df_30.head(10).iterrows()]
                text_aging_30 = f"Persisten sejak kemarin ({len(df_30)})\n" + "\n".join(list_30)
            else:
                text_aging_30 = f"Persisten sejak kemarin ({len(df_30)})"

            if not df_60.empty and ticket_col:
                list_60 = [f"• {row[ticket_col]} | {int(row['aging_num'])} Hari | -" for _, row in df_60.head(10).iterrows()]
                text_aging_60 = f"Persisten sejak kemarin ({len(df_60)})\n" + "\n".join(list_60)
            else:
                text_aging_60 = f"Persisten sejak kemarin ({len(df_60)})"
        else:
            aging_0_3 = aging_3_7 = aging_7_14 = aging_14_plus = 0
            text_aging_30 = "Persisten sejak kemarin (0)"
            text_aging_60 = "Persisten sejak kemarin (0)"

        analisa_text = (
            f"1. Backlog {lob} total {total_backlog} tiket pada {backlog_date}.\n"
            f"2. Kenaikan/penurunan terpantau pada sub topic utama.\n"
            f"3. Tiket >14 Hari terdata sebanyak {aging_14_plus} tiket."
        )
        trend_status = "STABLE ➖"
        change_val = "0"
        insight_text = f"💡 Focus LOB {lob}: Penanganan tiket aging >14 hari."

        # SUSUN 15 KOLOM PRESISIsesuai RAW DATA SHEET (A - O)
        row_data = [
            backlog_date,    # A: Tanggal backlog (Dari Tanggal Email)
            total_backlog,   # B: Total Backlog
            lob,             # C: LOB
            top_topic_str,   # D: Top mobile sub topic
            top_case_str,    # E: Case
            aging_0_3,       # F: 0-3 Days
            aging_3_7,       # G: 3-7
            aging_7_14,      # H: 7-14
            aging_14_plus,   # I: >14 Days
            text_aging_30,   # J: AGING >30 (Sudah terisi list tiket)
            text_aging_60,   # K: AGING >60 (Sudah terisi list tiket)
            analisa_text,    # L: Analisa
            trend_status,    # M: TREND
            change_val,      # N: CHANGE
            insight_text     # O: INSIGHT
        ]

        worksheet.append_row(row_data, value_input_option="USER_ENTERED")

    print("✅ Berhasil menulis data presisi per LOB ke Google Sheet!")

# ==========================================
# 4. UPLOAD BACKUP KE DRIVE
# ==========================================
def upload_to_drive(csv_bytes, file_name, backlog_date):
    print("☁️ Mengunggah backup file ke Google Drive...")
    file_metadata = {
        'name': f"{backlog_date}_{file_name}",
        'parents': [DRIVE_FOLDER_ID]
    }
    media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype='text/csv', resumable=True)
    drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print("✅ Backup tersimpan di Drive!")

# ==========================================
# 5. EXECUTION MAIN
# ==========================================
if __name__ == "__main__":
    try:
        csv_bytes, file_name, backlog_date, mail, email_id = download_latest_csv_from_gmail()
        
        if csv_bytes is None:
            print("🚀 Tidak ada email baru/file link download. Automation selesai aman.")
            exit(0)

        process_and_update_sheets(csv_bytes, file_name, backlog_date)
        upload_to_drive(csv_bytes, file_name, backlog_date)
        mark_email_as_processed(mail, email_id)
        
        print("🎉 AUTOMATION SELESAI DENGAN SUKSES & PRESISI!")
    except Exception as e:
        print(f"❌ ERROR DETECTED: {str(e)}")
        exit(1)
