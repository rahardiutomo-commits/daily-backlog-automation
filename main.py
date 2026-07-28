import os
import json
import requests
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# =====================================================
# AUTHENTICATION & CONFIG (GITHUB ACTIONS SAFE)
# =====================================================
SPREADSHEET_ID = "1oI-f_KPFqTwe8Q0M3zva1f2QbbsBeegDi-7Yly-W1Cs"
FOLDER_ID = "1zwQoKtOKf0houdAFIgC1MADFU8oygvt_"
CSV_URL = "https://storage.googleapis.com/kapture_report/Ticket_Report_sr_1467_1000221_1785199080906.csv"

# Ambil Service Account JSON dari Environment Variable GitHub Secrets
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds_json = os.environ.get("GCP_SA_KEY")
if not creds_json:
    raise ValueError("❌ Error: GCP_SA_KEY Secret tidak ditemukan di GitHub!")

service_account_info = json.loads(creds_json)
creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
gc = gspread.authorize(creds)

# =====================================================
# HELPER FUNCTIONS
# =====================================================
def top5(df, column):
    if column not in df.columns or df.empty:
        return ""
    counts = df[column].value_counts().head(5)
    return "\n".join([f"- {k}: {v}" for k, v in counts.items()])

def aging(df):
    age_col = "Days Between Create To Current Date"
    if age_col not in df.columns:
        return 0, 0, 0, 0

    age = pd.to_numeric(df[age_col], errors="coerce")
    return (
        len(age[age <= 3]),
        len(age[(age > 3) & (age <= 7)]),
        len(age[(age > 7) & (age <= 14)]),
        len(age[age > 14])
    )

# =====================================================
# DOWNLOAD CSV (NO 50MB LIMIT IN PYTHON)
# =====================================================
print("📥 Downloading CSV...")
r = requests.get(CSV_URL, stream=True)
with open("backlog.csv", "wb") as f:
    for chunk in r.iter_content(chunk_size=8192):
        f.write(chunk)
print("✅ Download Complete")

# =====================================================
# LOAD CSV & DETECT COLUMNS
# =====================================================
df = pd.read_csv("backlog.csv", low_memory=False)
print("Original Shape:", df.shape)

id_cols = [c for c in df.columns if "id" in str(c).lower() or "number" in str(c).lower() or "no" in str(c).lower()]
ticket_id_col = id_cols[0] if id_cols else df.columns[0]
print(f"🔑 Kolom Identifier Tiket: '{ticket_id_col}'")

# Hapus duplikat
initial_rows = len(df)
df = df.drop_duplicates(subset=[ticket_id_col], keep='first')
dropped_rows = initial_rows - len(df)
print(f"🧹 Pembersihan Duplikat: Menghapus {dropped_rows} tiket ganda. Sisa data: {len(df)} baris.")

df_original = df.copy()

# =====================================================
# ARCHIVE VERSION (NO PAYLOAD) & GOOGLE DRIVE UPLOAD
# =====================================================
payload_cols = [c for c in df.columns if "payload" in str(c).lower()]
archive_df = df.drop(columns=payload_cols, errors="ignore")

today = datetime.now().strftime("%Y-%m-%d")
archive_file = f"BACKLOG_{today}.csv"

archive_df.to_csv(archive_file, index=False)
print(f"✅ Local Archive Saved: {archive_file}")

try:
    print("📤 Uploading archive to Google Drive folder...")
    drive_service = build('drive', 'v3', credentials=creds)

    file_metadata = {
        'name': archive_file,
        'parents': [FOLDER_ID]
    }

    media = MediaFileUpload(archive_file, mimetype='text/csv', resumable=True)
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    print(f"🚀 SUCCESS! Backup Drive File ID: {uploaded_file.get('id')}")

except Exception as e:
    print(f"❌ Gagal upload ke Google Drive: {e}")

# =====================================================
# QUEUE MAPPING & SPLIT LOB
# =====================================================
queue_col = "Ticket Queue Name"

FRAUD_QUEUE = ["Bucket - CS Fraud", "CS Fraud", "L2 Fraud - Light Risk Fraud ATO", "View L3 fraud", "L2 Fraud - Tickets Escalation From CFM", "L2 Fraud - Tickets Escalation From Risk", "L2 Fraud - High Priority", "L2 Fraud - Tickets from Partner"]
NON_FRAUD_QUEUE = ["Agent Reset PIN L2", "Bucket - CS L2", "CS Support L2", "L2 Non Fraud - High Priority", "L2 Non Fraud - Tickets Escalation From TS Merchant & Channel", "L2 Non Fraud - DANA Bisnis", "L2 Non Fraud - Reset PIN (Re-Open Bulk Inactive Number)", "L2 Non Fraud - Change Number", "L2 Non Fraud - Tickets Escalation From Risk", "L2 Non Fraud - Tickets Escalation From AML", "L2 Non Fraud - Reset PIN (Re-Open Bulk Active Number)"]
CHANNEL_QUEUE = ["Channel Support", "Bucket - Channel Support", "L2 Channel - Quewise", "L2 Channel - Tickets From CFM", "L2 Channel - Tickets from Partner", "L2 Channel - All Ticket Over SLA"]
MERCHANT_QUEUE = ["Merchant Support", "Bucket - Merchant Support", "L2 Merchant - Tickets from Partner", "L2 Merchant - High Priority", "L2 Merchant - All QRIS Tickets", "L2 Merchant - Tickets Escalation From CFM", "L2 Merchant - Tickets Escalation From TS Merchant", "L2 Merchant - Quewise", "L2 Merchant - Tickets Escalation From Merchant Service", "QRIS_MS"]

fraud_df = df_original[df_original[queue_col].isin(FRAUD_QUEUE)]
nonfraud_df = df_original[df_original[queue_col].isin(NON_FRAUD_QUEUE)]
channel_df = df_original[df_original[queue_col].isin(CHANNEL_QUEUE)]
merchant_df = df_original[df_original[queue_col].isin(MERCHANT_QUEUE)]

# =====================================================
# OPEN GOOGLE SHEET & GET YESTERDAY DATE
# =====================================================
spreadsheet = gc.open_by_key(SPREADSHEET_ID)

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

for sheet_name, lob_name, lob_df in lob_targets:
    print(f"📊 Processing & Analyzing LOB: {lob_name}...")

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

    # Pilar 1
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

    # Pilar 2
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
    else:
        if yesterday_lob_df.empty:
            point2_text = "2. Kenaikan tertinggi Sub topic tidak terdeteksi (Data kemarin kosong)."

    # Pilar 3
    driver_text = "3. Tidak ada data case aktif hari ini."
    if not lob_df.empty and "Case" in lob_df.columns:
        case_counts = lob_df["Case"].value_counts()
        if not case_counts.empty:
            top_case_name = case_counts.index[0]
            top_case_volume = case_counts.iloc[0]
            driver_text = f"3. Kontributor backlog tertinggi utama hari ini adalah case '{top_case_name}' ({top_case_volume} tiket)."

    # Pilar 4
    age_col = "Days Between Create To Current Date"
    persistent_list = []
    new_list = []

    if age_col in lob_df.columns and not lob_df.empty:
        age_numeric_today = pd.to_numeric(lob_df[age_col], errors="coerce")
        today_over14_df = lob_df[age_numeric_today > 14]

        yesterday_over14_ids = set()
        if not yesterday_lob_df.empty and age_col in yesterday_lob_df.columns:
            age_numeric_yesterday = pd.to_numeric(yesterday_lob_df[age_col], errors="coerce")
            yesterday_over14_ids = set(yesterday_lob_df[age_numeric_yesterday > 14][ticket_id_col].dropna().astype(str))

        case_col_name = "Case" if "Case" in lob_df.columns else (sub_col if sub_col in lob_df.columns else lob_df.columns[0])

        for _, row in today_over14_df.iterrows():
            t_id = str(row[ticket_id_col])
            raw_age = pd.to_numeric(row[age_col], errors="coerce")
            t_age = str(int(raw_age)) if pd.notna(raw_age) else str(row[age_col])
            t_case = str(row[case_col_name]).strip() if pd.notna(row[case_col_name]) and str(row[case_col_name]).strip() != "" else "-"

            bullet_point = f"• {t_id} | {t_age} Hari | {t_case}"

            if t_id in yesterday_over14_ids:
                persistent_list.append(bullet_point)
            else:
                new_list.append(bullet_point)

    aging_text = "Clean! Tidak ada tiket >14 hari."
    if a3 > 0:
        persistent_str = "\n".join(persistent_list) if persistent_list else "• (Tidak ada)"
        new_str = "\n".join(new_list) if new_list else "• (Tidak ada)"

        aging_text = (
            f"⚠️ Ticket >14 Hari : {a3}\n\n"
            f"🟥 Persisten sejak kemarin ({len(persistent_list)})\n"
            f"{persistent_str}\n\n"
            f"🟨 Baru menjadi >14 Hari ({len(new_list)})\n"
            f"{new_str}"
        )

    full_analysis = f"1. {trend_text}\n{point2_text}\n{driver_text}\n\n{aging_text}"

    # Update LOB Sheet
    ws.clear()
    required_rows = len(lob_df) + 10
    required_cols = len(lob_df.columns) + 5
    if ws.row_count < required_rows: ws.add_rows(required_rows - ws.row_count)
    if ws.col_count < required_cols: ws.add_cols(required_cols - ws.col_count)

    ws.update(values=[lob_df.columns.tolist()], range_name="A1")
    chunk_size = 1000
    total = len(lob_df)
    for start in range(0, total, chunk_size):
        end = min(start + chunk_size, total)
        values = lob_df.iloc[start:end].fillna("").astype(str).values.tolist()
        ws.update(values=values, range_name=f"A{start+2}")

    print(f"✅ Sheet {sheet_name} Updated successfully.")

    history_rows_to_append.append([
        today,
        total_today,
        lob_name,
        top5(lob_df, "Mobile App - Sub Topic"),
        top5(lob_df, "Case"),
        a0,
        a1,
        a2,
        a3,
        full_analysis
    ])

# Append History
print("📝 Appending analytics to BACKLOG_HISTORY...")
history_ws.append_rows(history_rows_to_append)
print("\n🎉 DONE SUCCESSFULLY VIA GITHUB ACTIONS!")
