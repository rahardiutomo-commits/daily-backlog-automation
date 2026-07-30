import os
import io
import json
import imaplib
import email
import re
from datetime import datetime
from email.utils import parsedate_to_datetime
from bs4 import BeautifulSoup
import requests
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# =====================================================
# AUTHENTICATION & CONFIG
# =====================================================
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

SPREADSHEET_ID = "1oI-f_KPFqTwe8Q0M3zva1f2QbbsBeegDi-7Yly-W1Cs"

# =====================================================
# HELPER FUNCTIONS
# =====================================================
def top5(df, column):
    actual_col = column if column in df.columns else next((c for c in df.columns if column.lower() in str(c).lower()), None)
    if not actual_col or df.empty:
        return ""
    counts = df[actual_col].value_counts().head(5)
    return "\n".join([f"- {k}: {v}" for k, v in counts.items()])

def aging(df):
    age_col = "Days Between Create To Current Date"
    if age_col not in df.columns:
        age_col = next((c for c in df.columns if 'day' in str(c).lower() or 'aging' in str(c).lower()), None)
    
    if not age_col or age_col not in df.columns:
        return 0, 0, 0, 0

    age = pd.to_numeric(df[age_col], errors="coerce")
    return (
        len(age[age <= 3]),
        len(age[(age > 3) & (age <= 7)]),
        len(age[(age > 7) & (age <= 14)]),
        len(age[age > 14])
    )

# =====================================================
# DOWNLOAD CSV/EXCEL DARI GMAIL
# =====================================================
def download_latest_report_from_gmail():
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
        return None, None, mail, None

    latest_email_id = mail_ids[-1]
    status, msg_data = mail.fetch(latest_email_id, "(RFC822)")

    file_bytes = None
    email_date_str = None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            
            date_header = msg.get("Date")
            if date_header:
                try:
                    dt = parsedate_to_datetime(date_header)
                    email_date_str = dt.strftime('%Y-%m-%d')
                except Exception:
                    email_date_str = datetime.now().strftime('%Y-%m-%d')
            else:
                email_date_str = datetime.now().strftime('%Y-%m-%d')

            body_html = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/html':
                        body_html += part.get_payload(decode=True).decode('utf-8', errors='ignore')
            else:
                body_html = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

            soup = BeautifulSoup(body_html, 'html.parser')
            target_url = None

            for a_tag in soup.find_all('a', href=True):
                text = a_tag.get_text().lower()
                href = a_tag['href']
                if any(k in text for k in ['download', 'click', 'here', 'report']):
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
                    file_bytes = resp.content
                    print("✅ Berhasil mendownload report dari link!")

    return file_bytes, email_date_str, mail, latest_email_id

def mark_email_as_processed(mail, email_id):
    if not mail or not email_id:
        return
    try:
        mail.store(email_id, '+X-GM-LABELS', 'PROCESSED_BACKLOG')
        mail.store(email_id, '+FLAGS', '\\Seen')
        mail.logout()
        print("✅ Email berhasil ditandai!")
    except Exception as e:
        print(f"⚠️ Catatan pelabelan email: {str(e)}")

# =====================================================
# MAIN PROCESSING ENGINE (16 KOLOM PRESISI)
# =====================================================
def run_pipeline():
    file_bytes, today, mail, email_id = download_latest_report_from_gmail()
    if file_bytes is None:
        print("🚀 Selesai aman.")
        return

    # 1. LOAD CSV / EXCEL DENGAN MULTI-PARSER
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), low_memory=False)
    except Exception:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))
        except Exception:
            dfs = pd.read_html(io.BytesIO(file_bytes))
            df = dfs[0]

    df.columns = [str(c).strip() for c in df.columns]
    print("Original Shape:", df.shape)

    # Auto-detect kolom unik Ticket ID
    id_cols = [c for c in df.columns if "id" in str(c).lower() or "number" in str(c).lower() or "no" in str(c).lower()]
    ticket_id_col = id_cols[0] if id_cols else df.columns[0]
    print(f"🔑 Kolom Identifier Tiket yang terdeteksi: '{ticket_id_col}'")

    # --- PROSES HAPUS DUPLIKAT BERDASARKAN TICKET ID ---
    initial_rows = len(df)
    df = df.drop_duplicates(subset=[ticket_id_col], keep='first')
    dropped_rows = initial_rows - len(df)
    print(f"🧹 Pembersihan Duplikat: Berhasil menghapus {dropped_rows} tiket ganda. Sisa data: {len(df)} baris.")

    df_original = df.copy()

    # QUEUE MAPPING
    queue_col = "Ticket Queue Name"
    if queue_col not in df_original.columns:
        queue_col = next((c for c in df_original.columns if 'queue' in str(c).lower()), df_original.columns[0])

    FRAUD_QUEUE = ["Bucket - CS Fraud", "CS Fraud", "L2 Fraud - Light Risk Fraud ATO", "View L3 fraud", "L2 Fraud - Tickets Escalation From CFM", "L2 Fraud - Tickets Escalation From Risk", "L2 Fraud - High Priority", "L2 Fraud - Tickets from Partner"]
    NON_FRAUD_QUEUE = ["Agent Reset PIN L2", "Bucket - CS L2", "CS Support L2", "L2 Non Fraud - High Priority", "L2 Non Fraud - Tickets Escalation From TS Merchant & Channel", "L2 Non Fraud - DANA Bisnis", "L2 Non Fraud - Reset PIN (Re-Open Bulk Inactive Number)", "L2 Non Fraud - Change Number", "L2 Non Fraud - Tickets Escalation From Risk", "L2 Non Fraud - Tickets Escalation From AML", "L2 Non Fraud - Reset PIN (Re-Open Bulk Active Number)"]
    CHANNEL_QUEUE = ["Channel Support", "Bucket - Channel Support", "L2 Channel - Quewise", "L2 Channel - Tickets From CFM", "L2 Channel - Tickets from Partner", "L2 Channel - All Ticket Over SLA"]
    MERCHANT_QUEUE = ["Merchant Support", "Bucket - Merchant Support", "L2 Merchant - Tickets from Partner", "L2 Merchant - High Priority", "L2 Merchant - All QRIS Tickets", "L2 Merchant - Tickets Escalation From CFM", "L2 Merchant - Tickets Escalation From TS Merchant", "L2 Merchant - Quewise", "L2 Merchant - Tickets Escalation From Merchant Service", "QRIS_MS"]

    fraud_df = df_original[df_original[queue_col].isin(FRAUD_QUEUE)]
    nonfraud_df = df_original[df_original[queue_col].isin(NON_FRAUD_QUEUE)]
    channel_df = df_original[df_original[queue_col].isin(CHANNEL_QUEUE)]
    merchant_df = df_original[df_original[queue_col].isin(MERCHANT_QUEUE)]

    # OPEN GOOGLE SHEET & AMBIL HEADER BACKLOG_HISTORY ASLI
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    history_ws = spreadsheet.worksheet("BACKLOG_HISTORY")
    history_data = history_ws.get_all_values()
    
    # Header Sheet dijadikan Acuan Kunci Posisi Kolom
    sheet_headers = [str(h).strip() for h in history_data[0]] if history_data else []
    
    yesterday = "kemarin"
    if len(history_data) > 1:
        history_df = pd.DataFrame(history_data[1:], columns=sheet_headers)
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

    history_rows_to_append = []

    lob_targets = [
        ("FRAUD", "Fraud", fraud_df),
        ("NON FRAUD", "Non Fraud", nonfraud_df),
        ("CHANNEL", "Channel", channel_df),
        ("MERCHANT", "Merchant", merchant_df)
    ]

    for sheet_name, lob_name, lob_df in lob_targets:
        print(f"📊 Processing & Analyzing LOB: {lob_name} (Total: {len(lob_df)} tiket)...")

        yesterday_lob_df = pd.DataFrame()
        try:
            ws = spreadsheet.worksheet(sheet_name)
            old_data = ws.get_all_values()

            if old_data and len(old_data) > 1:
                headers = [str(h).strip() for h in old_data[0]]
                seen = {}
                cleaned_headers = []
                for i, h in enumerate(headers):
                    if h == "": h = f"UnnamedColumn_{i}"
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

        # 1. PILAR ANALISA TREND & CHANGE
        trend_status = "STABLE ➖"
        change_val = "0"
        trend_text = "Data hari pertama (belum ada pembanding)."
        if total_yesterday > 0:
            diff = total_today - total_yesterday
            pct = round(abs(diff) / total_yesterday * 100, 1) if total_yesterday else 0
            change_val = f"{diff:+,}"
            if diff > 0:
                trend_status = f"UP ⬆️ (+{pct}%)"
                trend_text = f"Backlog naik {pct}% ({diff:+,}) dibanding tanggal {yesterday}."
            elif diff < 0:
                trend_status = f"DOWN ⬇️ (-{pct}%)"
                trend_text = f"Backlog turun {pct}% ({diff:+,}) dibanding tanggal {yesterday}."
            else:
                trend_text = f"Backlog stabil dibanding tanggal {yesterday}."

        # 2. PILAR KENAIKAN SUB TOPIC
        sub_col = "Mobile App - Sub Topic"
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

        # 3. PILAR KONTRIBUTOR CASE
        driver_text = "3. Tidak ada data case aktif hari ini."
        if not lob_df.empty and "Case" in lob_df.columns:
            case_counts = lob_df["Case"].value_counts()
            if not case_counts.empty:
                driver_text = f"3. Kontributor backlog tertinggi utama hari ini adalah case '{case_counts.index[0]}' ({case_counts.iloc[0]} tiket)."

        # 4. BREAKDOWN AGING >14, >30, >60
        age_col = "Days Between Create To Current Date"
        if age_col not in lob_df.columns:
            age_col = next((c for c in lob_df.columns if 'day' in str(c).lower() or 'aging' in str(c).lower()), None)

        persistent_list, new_list, aging_30_list, aging_60_list = [], [], [], []

        if age_col and age_col in lob_df.columns and not lob_df.empty:
            age_numeric_today = pd.to_numeric(lob_df[age_col], errors="coerce").fillna(0)
            
            today_over14_df = lob_df[age_numeric_today > 14]
            today_over30_df = lob_df[age_numeric_today > 30]
            today_over60_df = lob_df[age_numeric_today > 60]

            yesterday_over14_ids = set()
            if not yesterday_lob_df.empty and age_col in yesterday_lob_df.columns:
                age_numeric_yesterday = pd.to_numeric(yesterday_lob_df[age_col], errors="coerce").fillna(0)
                yesterday_over14_ids = set(yesterday_lob_df[age_numeric_yesterday > 14][ticket_id_col].dropna().astype(str))

            case_col_name = "Case" if "Case" in lob_df.columns else (sub_col if sub_col in lob_df.columns else lob_df.columns[0])

            # AGING > 14
            for _, row in today_over14_df.iterrows():
                t_id = str(row[ticket_id_col])
                raw_age = pd.to_numeric(row[age_col], errors="coerce")
                t_age = str(int(raw_age)) if pd.notna(raw_age) else str(row[age_col])
                t_case = str(row[case_col_name]).strip() if pd.notna(row[case_col_name]) and str(row[case_col_name]).strip() != "" else "-"
                bullet = f"• {t_id} | {t_age} Hari | {t_case}"
                if t_id in yesterday_over14_ids:
                    persistent_list.append(bullet)
                else:
                    new_list.append(bullet)

            # AGING > 30
            for _, row in today_over30_df.head(10).iterrows():
                aging_30_list.append(f"• {str(row[ticket_id_col])} | {int(pd.to_numeric(row[age_col], errors='coerce'))} Hari")

            # AGING > 60
            for _, row in today_over60_df.head(10).iterrows():
                aging_60_list.append(f"• {str(row[ticket_id_col])} | {int(pd.to_numeric(row[age_col], errors='coerce'))} Hari")

        aging_text = "Clean! Tidak ada tiket >14 hari."
        if a3 > 0:
            p_str = "\n".join(persistent_list) if persistent_list else "• (Tidak ada)"
            n_str = "\n".join(new_list) if new_list else "• (Tidak ada)"
            aging_text = f"⚠️ Ticket >14 Hari : {a3}\n\n🟥 Persisten sejak kemarin ({len(persistent_list)})\n{p_str}\n\n🟨 Baru menjadi >14 Hari ({len(new_list)})\n{n_str}"

        full_analysis = f"1. {trend_text}\n{point2_text}\n{driver_text}\n\n{aging_text}"
        
        text_aging_30 = f"Persisten ({len(aging_30_list)})\n" + "\n".join(aging_30_list) if aging_30_list else f"Persisten ({len(aging_30_list)})"
        text_aging_60 = f"Persisten ({len(aging_60_list)})\n" + "\n".join(aging_60_list) if aging_60_list else f"Persisten ({len(aging_60_list)})"
        insight_text = f"💡 Focus LOB {lob_name}: Penanganan tiket aging >14 hari."

        # UPDATE TAB LOB INDIVIDU
        try:
            ws = spreadsheet.worksheet(sheet_name)
            ws.clear()
            req_r, req_c = len(lob_df) + 10, len(lob_df.columns) + 5
            if ws.row_count < req_r: ws.add_rows(req_r - ws.row_count)
            if ws.col_count < req_c: ws.add_cols(req_c - ws.col_count)

            ws.update(values=[lob_df.columns.tolist()], range_name="A1")
            chunk_size = 1000
            for start in range(0, len(lob_df), chunk_size):
                end = min(start + chunk_size, len(lob_df))
                ws.update(values=lob_df.iloc[start:end].fillna("").astype(str).values.tolist(), range_name=f"A{start+2}")

            print(f"✅ Sheet {sheet_name} Updated successfully.")
        except Exception as e:
            print(f"⚠️ Warning tab {sheet_name}: {e}")

        # DICTIONARY PENETAPAN 16 KOLOM KUNCI SAMA DENGAN HEADER SHEET
        row_dict = {
            "Tanggal backlog": today,
            "Total Backlog": total_today,
            "LOB (Fraud/NonFraud/Merchant/channel)": lob_name,
            "LOB": lob_name,
            "Top mobile sub topic": top5(lob_df, "Mobile App - Sub Topic"),
            "Case": top5(lob_df, "Case"),
            "Aging ticket 0-3 Days": a0,
            "3-7": a1,
            "7-14": a2,
            ">14": a3,
            "AGING >30": text_aging_30,
            "AGING >60": text_aging_60,
            "Analisa": full_analysis,
            "TREND": trend_status,
            "CHANGE": change_val,
            "INSIGHT": insight_text
        }

        # BUAT BARIS DATA PRESISI MENGIKUTI STRUKTUR HEADER SHEET (A s.d P / 16 KOLOM)
        final_row = []
        for header in sheet_headers:
            val = row_dict.get(header.strip(), "")
            final_row.append(val)

        history_rows_to_append.append(final_row)

    print("📝 Appending 16 columns using header mapping to BACKLOG_HISTORY...")
    history_ws.append_rows(history_rows_to_append, value_input_option="USER_ENTERED")
    mark_email_as_processed(mail, email_id)
    print("\n🎉 DONE SUCCESSFULLY WITH COMPLETE 16-COLUMN ANALYSIS!")

if __name__ == "__main__":
    run_pipeline()
