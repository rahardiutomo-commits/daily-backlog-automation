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
TARGET_SUBJECT = "All Pending - Backlog RTFM [06:00 AM]"

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
# GMAIL SWEEPING FUNCTIONS
# =====================================================
def get_unprocessed_emails():
    """Mengambil SEMUA email yang memiliki subject spesifik dan belum dilabeli PROCESSED_BACKLOG."""
    print(f"📧 Memeriksa email masuk dengan subject '{TARGET_SUBJECT}'...")
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    # Search spesifik subject dan belum dilabeli
    search_query = f'SUBJECT "{TARGET_SUBJECT}" NOT label:PROCESSED_BACKLOG'
    status, messages = mail.search(None, 'X-GM-RAW', search_query)

    mail_ids = messages[0].split()
    print(f"📥 Ditemukan {len(mail_ids)} email yang perlu diproses.")
    return mail, mail_ids

def extract_email_data(mail, email_id):
    """Mengekstrak TANGGAL EMAIL dan FILE REPORT secara presisi."""
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    file_bytes = None
    email_date_str = None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            
            # --- EXTRACT TANGGAL ASLI EMAIL ---
            date_header = msg.get("Date")
            if date_header:
                try:
                    dt = parsedate_to_datetime(date_header)
                    email_date_str = dt.strftime('%Y-%m-%d')
                except Exception as e:
                    print(f"⚠️ Warning parse header date: {e}")
                    email_date_str = datetime.now().strftime('%Y-%m-%d')
            else:
                email_date_str = datetime.now().strftime('%Y-%m-%d')

            # Ekstrak Body HTML
            body_html = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/html':
                        body_html += part.get_payload(decode=True).decode('utf-8', errors='ignore')
            else:
                body_html = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

            # Cari Link Download Report
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
                print(f"🔗 Link download ditemukan (Tanggal Email: {email_date_str}): {target_url}")
                resp = requests.get(target_url, timeout=60)
                if resp.status_code == 200:
                    file_bytes = resp.content
                    print("✅ Berhasil mendownload file report!")

    return file_bytes, email_date_str

def mark_email_as_processed(mail, email_id):
    """Tandai email dengan label PROCESSED_BACKLOG agar tidak di-sweep ulang."""
    if not mail or not email_id:
        return
    try:
        mail.store(email_id, '+X-GM-LABELS', 'PROCESSED_BACKLOG')
        mail.store(email_id, '+FLAGS', '\\Seen')
        print(f"🏷️ Email ID {email_id.decode()} berhasil ditandai sebagai 'PROCESSED_BACKLOG'!")
    except Exception as e:
        print(f"⚠️ Catatan pelabelan email ID {email_id}: {str(e)}")

# =====================================================
# MAIN PROCESSING ENGINE
# =====================================================
def run_pipeline():
    mail, mail_ids = get_unprocessed_emails()
    
    if not mail_ids:
        print("ℹ️ Tidak ada email baru yang sesuai kriteria hari ini.")
        mail.logout()
        return

    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    history_ws = spreadsheet.worksheet("BACKLOG_HISTORY")

    # SWEEPING PROSES SETIAP EMAIL 1 PER 1
    for index, email_id in enumerate(mail_ids, start=1):
        file_bytes, email_date = extract_email_data(mail, email_id)
        
        print(f"\n==================================================")
        print(f"🔄 Memproses Email {index} dari {len(mail_ids)} | Tanggal Email: {email_date}")
        print(f"==================================================")

        if file_bytes is None:
            print(f"⚠️ Gagal mengunduh attachment dari email ID {email_id.decode()}, melewatinya.")
            continue

        # 1. LOAD CSV / EXCEL DATA
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), low_memory=False)
        except Exception:
            try:
                df = pd.read_excel(io.BytesIO(file_bytes))
            except Exception:
                dfs = pd.read_html(io.BytesIO(file_bytes))
                df = dfs[0]

        df.columns = [str(c).strip() for c in df.columns]
        print(f"📄 Shape Data: {df.shape}")

        id_cols = [c for c in df.columns if "id" in str(c).lower() or "number" in str(c).lower() or "no" in str(c).lower()]
        ticket_id_col = id_cols[0] if id_cols else df.columns[0]

        # Cleansing Duplikat Tiket
        initial_rows = len(df)
        df = df.drop_duplicates(subset=[ticket_id_col], keep='first')
        print(f"🧹 Dihapus {initial_rows - len(df)} tiket ganda. Sisa data unik: {len(df)} baris.")

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

        # BACA HISTORIS DARI SHEET UNTUK HITUNG DIFFERENCE/CHANGE
        history_data = history_ws.get_all_values()
        sheet_headers = [str(h).strip() for h in history_data[0]] if history_data else []
        
        yesterday = "kemarin"
        if len(history_data) > 1:
            history_df = pd.DataFrame(history_data[1:], columns=sheet_headers)
            if "Tanggal backlog" in history_df.columns:
                try:
                    raw_dates = history_df[history_df["Tanggal backlog"].str.strip() != ""]["Tanggal backlog"].unique()
                    parsed_dates = pd.to_datetime(raw_dates, errors='coerce')
                    email_dt = pd.to_datetime(email_date)
                    past_dates = [d.strftime("%Y-%m-%d") for d in parsed_dates if pd.notna(d) and d < email_dt]
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
            print(f"📊 Analisa LOB: {lob_name} ({len(lob_df)} tiket)...")

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
                print(f"ℹ️ Selisih historis tab {sheet_name}: {e}")

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

            # 2. SUB TOPIC & 3. CASE DRIVER
            sub_col = "Mobile App - Sub Topic"
            point2_text = "2. Data Sub topic tidak tersedia untuk pembanding."
            if sub_col in lob_df.columns and not yesterday_lob_df.empty and sub_col in yesterday_lob_df.columns:
                sub_diff = lob_df[sub_col].value_counts().sub(yesterday_lob_df[sub_col].value_counts(), fill_value=0)
                if not sub_diff.empty and sub_diff.max() > 0:
                    point2_text = f"2. Kenaikan tertinggi pada Sub topic '{sub_diff.idxmax()}' (+{int(sub_diff.max())} tiket)."
                else:
                    point2_text = "2. Tidak ada kenaikan volume Sub topic dibanding kemarin."

            driver_text = "3. Tidak ada data case aktif hari ini."
            if not lob_df.empty and "Case" in lob_df.columns:
                case_counts = lob_df["Case"].value_counts()
                if not case_counts.empty:
                    driver_text = f"3. Kontributor backlog tertinggi hari ini adalah case '{case_counts.index[0]}' ({case_counts.iloc[0]} tiket)."

            # 4. AGING BREAKDOWN
            age_col = "Days Between Create To Current Date"
            if age_col not in lob_df.columns:
                age_col = next((c for c in lob_df.columns if 'day' in str(c).lower() or 'aging' in str(c).lower()), None)

            persistent_list, new_list, aging_30_list, aging_60_list = [], [], [], []

            if age_col and age_col in lob_df.columns and not lob_df.empty:
                age_num = pd.to_numeric(lob_df[age_col], errors="coerce").fillna(0)
                today_over14 = lob_df[age_num > 14]
                
                yesterday_over14_ids = set()
                if not yesterday_lob_df.empty and age_col in yesterday_lob_df.columns:
                    y_age_num = pd.to_numeric(yesterday_lob_df[age_col], errors="coerce").fillna(0)
                    yesterday_over14_ids = set(yesterday_lob_df[y_age_num > 14][ticket_id_col].dropna().astype(str))

                case_col_name = "Case" if "Case" in lob_df.columns else (sub_col if sub_col in lob_df.columns else lob_df.columns[0])

                for _, row in today_over14.iterrows():
                    t_id = str(row[ticket_id_col])
                    t_age = str(int(pd.to_numeric(row[age_col], errors="coerce"))) if pd.notna(pd.to_numeric(row[age_col], errors="coerce")) else str(row[age_col])
                    t_case = str(row[case_col_name]).strip() if pd.notna(row[case_col_name]) else "-"
                    bullet = f"• {t_id} | {t_age} Hari | {t_case}"
                    if t_id in yesterday_over14_ids:
                        persistent_list.append(bullet)
                    else:
                        new_list.append(bullet)

                for _, row in lob_df[age_num > 30].head(10).iterrows():
                    aging_30_list.append(f"• {str(row[ticket_id_col])} | {int(pd.to_numeric(row[age_col], errors='coerce'))} Hari")

                for _, row in lob_df[age_num > 60].head(10).iterrows():
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

            # UPDATE TAB INDIVIDU LOB
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
                print(f"✅ Sheet {sheet_name} Updated.")
            except Exception as e:
                print(f"⚠️ Warning tab {sheet_name}: {e}")

            # --- MAPPING DATA DENGAN TANGGAL EMAIL DINAMIS (`email_date`) ---
            row_dict = {
                "Tanggal backlog": email_date,  # <-- Menggunakan tanggal asli email
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

            final_row = [row_dict.get(h.strip(), "") for h in sheet_headers]
            history_rows_to_append.append(final_row)

        print(f"📝 Menyimpan hasil analisa tanggal {email_date} ke BACKLOG_HISTORY...")
        history_ws.append_rows(history_rows_to_append, value_input_option="USER_ENTERED")

        # TANDAI EMAIL SELESAI
        mark_email_as_processed(mail, email_id)

    mail.logout()
    print("\n🎉 SELURUH EMAIL BERHASIL DISAPU DAN DIPROSES DENGAN TANGGAL MASING-MASING!")

if __name__ == "__main__":
    run_pipeline()
