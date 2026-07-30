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

# ==========================================
# 1. KREDENSIAL DARI ENVIRONMENT
# ==========================================
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
GCP_SA_KEY = os.environ.get("GCP_SA_KEY")

if not all([GMAIL_USER, GMAIL_APP_PASSWORD, GCP_SA_KEY]):
    raise ValueError("⚠️ Secret GitHub belum lengkap! Cek GMAIL_USER, GMAIL_APP_PASSWORD, dan GCP_SA_KEY.")

sa_info = json.loads(GCP_SA_KEY)
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = Credentials.from_service_account_info(sa_info, scopes=scopes)
gc = gspread.authorize(credentials)

SPREADSHEET_ID = "1oI-f_KPFqTwe8Q0M3zva1f2QbbsBeegDi-7Yly-W1Cs"

# ==========================================
# HELPER FUNCTIONS (AGING & TOP5)
# ==========================================
def aging(df):
    age_col = "Days Between Create To Current Date"
    if age_col not in df.columns or df.empty:
        return 0, 0, 0, 0
    
    age_numeric = pd.to_numeric(df[age_col], errors="coerce").fillna(0)
    a0 = len(df[age_numeric <= 3])
    a1 = len(df[(age_numeric > 3) & (age_numeric <= 7)])
    a2 = len(df[(age_numeric > 7) & (age_numeric <= 14)])
    a3 = len(df[age_numeric > 14])
    return a0, a1, a2, a3

def top5(df, col_name):
    if col_name not in df.columns or df.empty:
        return ""
    counts = df[col_name].value_counts().head(5)
    return "\n".join([f"- {k}: {v}" for k, v in counts.items()])

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
    file_name = "backlog_report.xlsx"
    email_date_str = None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            
            # --- TANGGAL BACKLOG DARI TANGGAL EMAIL ---
            date_header = msg.get("Date")
            if date_header:
                try:
                    dt = parsedate_to_datetime(date_header)
                    email_date_str = dt.strftime('%Y-%m-%d')
                except Exception:
                    email_date_str = pd.Timestamp.now().strftime('%Y-%m-%d')
            else:
                email_date_str = pd.Timestamp.now().strftime('%Y-%m-%d')

            # --- EKSTRAKSI LINK DOWNLOAD HREF ---
            body_html = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == 'text/html':
                        body_html += part.get_payload(decode=True).decode('utf-8', errors='ignore')
            else:
                body_html = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

            soup = BeautifulSoup(body_html, 'html.parser')
            target_url = None

            for a_tag in soup.find_all('a', href=True):
                text = a_tag.get_text().lower()
                href = a_tag['href']
                if 'download' in text or 'click' in text or 'here' in text or 'report' in text:
                    target_url = href
                    break
            
            if not target_url:
                urls = re.findall(r'https?://[^\s<>"]+', body_html)
                if urls:
                    target_url = urls[0]

            if target_url:
                print(f"🔗 Link download ditemukan: {target_url}")
                resp = requests.get(target_url, timeout=60)
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
        mail.store(email_id, '+X-GM-LABELS', 'PROCESSED_BACKLOG')
        mail.store(email_id, '+FLAGS', '\\Seen')
        mail.logout()
        print("✅ Email berhasil diberi label & ditandai terbaca!")
    except Exception as e:
        print(f"⚠️ Catatan pelabelan email: {str(e)}")

# ==========================================
# 3. PROSES DATA & FULL ANALYTICS
# ==========================================
def process_and_update_sheets(csv_bytes, file_name, today):
    print(f"📊 Menganalisis data untuk Tanggal Backlog: {today}...")
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    try:
        df_original = pd.read_excel(io.BytesIO(csv_bytes))
    except Exception:
        df_original = pd.read_csv(io.BytesIO(csv_bytes))

    df_original.columns = [str(c).strip() for c in df_original.columns]

    # QUEUE MAPPING & SPLIT LOB
    queue_col = "Ticket Queue Name"
    if queue_col not in df_original.columns:
        queue_col = next((c for c in df_original.columns if 'queue' in c.lower()), queue_col)

    FRAUD_QUEUE = ["Bucket - CS Fraud", "CS Fraud", "L2 Fraud - Light Risk Fraud ATO", "View L3 fraud", "L2 Fraud - Tickets Escalation From CFM", "L2 Fraud - Tickets Escalation From Risk", "L2 Fraud - High Priority", "L2 Fraud - Tickets from Partner"]
    NON_FRAUD_QUEUE = ["Agent Reset PIN L2", "Bucket - CS L2", "CS Support L2", "L2 Non Fraud - High Priority", "L2 Non Fraud - Tickets Escalation From TS Merchant & Channel", "L2 Non Fraud - DANA Bisnis", "L2 Non Fraud - Reset PIN (Re-Open Bulk Inactive Number)", "L2 Non Fraud - Change Number", "L2 Non Fraud - Tickets Escalation From Risk", "L2 Non Fraud - Tickets Escalation From AML", "L2 Non Fraud - Reset PIN (Re-Open Bulk Active Number)"]
    CHANNEL_QUEUE = ["Channel Support", "Bucket - Channel Support", "L2 Channel - Quewise", "L2 Channel - Tickets From CFM", "L2 Channel - Tickets from Partner", "L2 Channel - All Ticket Over SLA"]
    MERCHANT_QUEUE = ["Merchant Support", "Bucket - Merchant Support", "L2 Merchant - Tickets from Partner", "L2 Merchant - High Priority", "L2 Merchant - All QRIS Tickets", "L2 Merchant - Tickets Escalation From CFM", "L2 Merchant - Tickets Escalation From TS Merchant", "L2 Merchant - Quewise", "L2 Merchant - Tickets Escalation From Merchant Service", "QRIS_MS"]

    fraud_df = df_original[df_original[queue_col].isin(FRAUD_QUEUE)]
    nonfraud_df = df_original[df_original[queue_col].isin(NON_FRAUD_QUEUE)]
    channel_df = df_original[df_original[queue_col].isin(CHANNEL_QUEUE)]
    merchant_df = df_original[df_original[queue_col].isin(MERCHANT_QUEUE)]

    # =====================================================
    # OPEN GOOGLE SHEET & GET YESTERDAY DATE (FIXED)
    # =====================================================
    history_ws = spreadsheet.worksheet("BACKLOG_HISTORY")
    history_data = history_ws.get_all_values()
    yesterday = "kemarin"

    if len(history_data) > 1:
        headers = [str(h).strip() for h in history_data[0]]
        history_df = pd.DataFrame(history_data[1:], columns=headers)

        if "Tanggal backlog" in history_df.columns:
            try:
                raw_dates = history_df[history_df["Tanggal backlog"].str.strip() != ""]["Tanggal backlog"].unique()
                parsed_dates = pd.to_datetime(raw_dates, errors='coerce')
                today_dt = pd.to_datetime(today)
                past_dates = [d.strftime("%Y-%m-%d") for d in parsed_dates if pd.notna(d) and d < today_dt]

                if len(past_dates) > 0:
                    yesterday = max(past_dates)
            except Exception as e:
                print("⚠️ Gagal membaca tanggal kemarin dari history:", e)

    # =====================================================
    # PROCESS LOB, ANALYSIS, & UPLOAD TARGETS
    # =====================================================
    history_rows_to_append = []

    lob_targets = [
        ("FRAUD", "Fraud", fraud_df),
        ("NON FRAUD", "Non Fraud", nonfraud_df),
        ("CHANNEL", "Channel", channel_df),
        ("MERCHANT", "Merchant", merchant_df)
    ]

    ticket_id_col = next((c for c in df_original.columns if 'ticket' in c.lower() or 'id' in c.lower() or 'number' in c.lower()), df_original.columns[0])

    for sheet_name, lob_name, lob_df in lob_targets:
        print(f"📊 Processing & Analyzing LOB: {lob_name}...")

        # STEP A: AMBIL DATA KEMARIN MENGGUNAKAN GET_ALL_VALUES
        yesterday_lob_df = pd.DataFrame()
        try:
            ws = spreadsheet.worksheet(sheet_name)
            old_data = ws.get_all_values()

            if old_data and len(old_data) > 1:
                headers = [str(h).strip() for h in old_data[0]]
                seen = {}
                cleaned_headers = []
                for i, h in enumerate(headers):
                    if h == "":
                        h = f"UnnamedColumn_{i}"
                    if h in seen:
                        seen[h] += 1
                        cleaned_headers.append(f"{h}.{seen[h]}")
                    else:
                        seen[h] = 0
                        cleaned_headers.append(h)

                yesterday_lob_df = pd.DataFrame(old_data[1:], columns=cleaned_headers)
        except Exception as e:
            print(f"ℹ️ Gagal mengambil selisih historis untuk tab {sheet_name}: {e}")

        total_today = len(lob_df)
        total_yesterday = len(yesterday_lob_df) if not yesterday_lob_df.empty else 0
        a0, a1, a2, a3 = aging(lob_df)

        # PILAR 1: ANALISA TREND VOLUME
        trend_text = "Data hari pertama (belum ada pembanding)."
        if total_yesterday > 0:
            diff = total_today - total_yesterday
            pct = round(abs(diff) / total_yesterday * 100, 1) if total_yesterday else 0
            if diff > 0:
                trend_text = f"Backlog naik {pct}% ({diff:+,}) dibanding tanggal {yesterday}."
            elif diff < 0:
                trend_text = f"Backlog turun {pct}% ({diff:+,}) dibanding tanggal {yesterday}."
            else:
                trend_text = f"Backlog stabil dibanding tanggal {yesterday}."

        # PILAR 2: KENAIKAN TERTINGGI SUB TOPIC
        sub_col = "Mobile App - Sub Topic"
        if sub_col not in lob_df.columns:
            sub_col = next((c for c in lob_df.columns if 'sub' in c.lower() or 'topic' in c.lower()), "Sub Topic")

        point2_text = "2. Data Sub topic tidak tersedia untuk pembanding."
        if sub_col in lob_df.columns and not yesterday_lob_df.empty and sub_col in yesterday_lob_df.columns:
            today_subs = lob_df[sub_col].value_counts()
            yesterday_subs = yesterday_lob_df[sub_col].value_counts()

            sub_diff = today_subs.sub(yesterday_subs, fill_value=0)
            if not sub_diff.empty:
                max_inc_sub = sub_diff.idxmax()
                max_inc_val = sub_diff.max()
                if max_inc_val > 0:
                    point2_text = f"2. Kenaikan tertinggi terjadi pada Sub topic '{max_inc_sub}' (+{int(max_inc_val)} tiket)."
                else:
                    point2_text = "2. Tidak ada kenaikan volume pada Sub topic mana pun dibanding kemarin."
        else:
            if yesterday_lob_df.empty:
                point2_text = "2. Kenaikan tertinggi Sub topic tidak terdeteksi (Data kemarin kosong)."

        # PILAR 3: KONTRIBUTOR UTAMA CASE
        driver_text = "3. Tidak ada data case aktif hari ini."
        case_col_name = "Case" if "Case" in lob_df.columns else next((c for c in lob_df.columns if 'case' in c.lower()), None)
        if case_col_name and not lob_df.empty:
            case_counts = lob_df[case_col_name].value_counts()
            if not case_counts.empty:
                top_case_name = case_counts.index[0]
                top_case_volume = case_counts.iloc[0]
                driver_text = f"3. Kontributor backlog tertinggi utama hari ini adalah case '{top_case_name}' ({top_case_volume} tiket)."

        # PILAR 4: BREAKDOWN DETAIL TICKETS >14 HARI (PERSISTEN VS BARU)
        age_col = "Days Between Create To Current Date"
        if age_col not in lob_df.columns:
            age_col = next((c for c in lob_df.columns if 'aging' in c.lower() or 'day' in c.lower()), "Days")

        persistent_list = []
        new_list = []

        if age_col in lob_df.columns and not lob_df.empty:
            age_numeric_today = pd.to_numeric(lob_df[age_col], errors="coerce")
            today_over14_df = lob_df[age_numeric_today > 14]

            yesterday_over14_ids = set()
            if not yesterday_lob_df.empty and age_col in yesterday_lob_df.columns:
                age_numeric_yesterday = pd.to_numeric(yesterday_lob_df[age_col], errors="coerce")
                yesterday_over14_ids = set(yesterday_lob_df[age_numeric_yesterday > 14][ticket_id_col].dropna().astype(str))

            c_name = case_col_name if case_col_name else (sub_col if sub_col in lob_df.columns else lob_df.columns[0])

            for _, row in today_over14_df.iterrows():
                t_id = str(row[ticket_id_col])
                raw_age = pd.to_numeric(row[age_col], errors="coerce")
                t_age = str(int(raw_age)) if pd.notna(raw_age) else str(row[age_col])
                t_case = str(row[c_name]).strip() if pd.notna(row[c_name]) and str(row[c_name]).strip() != "" else "-"

                bullet_point = f"• {t_id} | {t_age} Hari | {t_case}"

                if t_id in yesterday_over14_ids:
                    persistent_list.append(bullet_point)
                else:
                    new_list.append(bullet_point)

        aging_text = "Clean! Tidak ada tiket >14 hari."
        if a3 > 0:
            persistent_str = "\n".join(persistent_list[:10]) if persistent_list else "• (Tidak ada)"
            new_str = "\n".join(new_list[:10]) if new_list else "• (Tidak ada)"

            aging_text = (
                f"⚠️ Ticket >14 Hari : {a3}\n\n"
                f"🟥 Persisten sejak kemarin ({len(persistent_list)})\n"
                f"{persistent_str}\n\n"
                f"🟨 Baru menjadi >14 Hari ({len(new_list)})\n"
                f"{new_str}"
            )

        # GABUNGKAN FOUR-PILLARS ANALYSIS
        full_analysis = f"1. {trend_text}\n{point2_text}\n{driver_text}\n\n{aging_text}"

        # STEP B: UPLOAD DATA BARU KE TAB LOB (FRAUD, NON FRAUD, CHANNEL, MERCHANT)
        try:
            ws_target = spreadsheet.worksheet(sheet_name)
            ws_target.clear()
            required_rows = len(lob_df) + 10
            required_cols = len(lob_df.columns) + 5
            if ws_target.row_count < required_rows: ws_target.add_rows(required_rows - ws_target.row_count)
            if ws_target.col_count < required_cols: ws_target.add_cols(required_cols - ws_target.col_count)

            ws_target.update(values=[lob_df.columns.tolist()], range_name="A1")
            chunk_size = 1000
            total = len(lob_df)
            for start in range(0, total, chunk_size):
                end = min(start + chunk_size, total)
                values = lob_df.iloc[start:end].fillna("").astype(str).values.tolist()
                ws_target.update(values=values, range_name=f"A{start+2}")

            print(f"✅ Sheet {sheet_name} Updated successfully.")
        except Exception as e:
            print(f"⚠️ Warning updating tab {sheet_name}: {e}")

        # Simpan baris analisa ke array history
        history_rows_to_append.append([
            today,
            total_today,
            lob_name,
            top5(lob_df, sub_col),
            top5(lob_df, case_col_name if case_col_name else "Case"),
            a0,
            a1,
            a2,
            a3,
            full_analysis
        ])

    # APPEND TO BACKLOG_HISTORY
    print("📝 Appending all analytics to BACKLOG_HISTORY...")
    history_ws.append_rows(history_rows_to_append, value_input_option="USER_ENTERED")
    print("🎉 DONE SUCCESSFULLY WITH ADVANCED COMPARATIVE ANALYSIS!")

# ==========================================
# 4. EXECUTION MAIN
# ==========================================
if __name__ == "__main__":
    try:
        csv_bytes, file_name, backlog_date, mail, email_id = download_latest_csv_from_gmail()
        
        if csv_bytes is None:
            print("🚀 Tidak ada email baru/file link download. Automation selesai aman.")
            exit(0)

        process_and_update_sheets(csv_bytes, file_name, backlog_date)
        mark_email_as_processed(mail, email_id)
        
        print("🎉 AUTOMATION SELESAI DENGAN SUKSES SECARA KESELURUHAN!")
    except Exception as e:
        print(f"❌ ERROR DETECTED: {str(e)}")
        exit(1)
