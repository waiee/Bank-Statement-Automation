import os
import pandas as pd
import re

DATA_DIR = "data"
OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "WC_statements.xlsx")
YEAR = "2025"

TEMPLATE_COLUMNS = [
    "DocNo", "DocNo2", "DocDate", "TaxDate", "DocType", "JournalType",
    "DealWith", "TaxEntity", "Description", "Extracted Description",
    "CurrencyCode", "PaymentAmt", "BankCharge", "BankChargeTaxCode",
    "BankChargeTaxRate", "BankChargeTax", "ToBankRate", "ToAccountRate",
    "BankChargeTaxRefNo", "PaymentBy", "FloatDay", "BankChargeDeptNo"
]

MONTH_MAP = {
    "JAN": 1, "JANUARY": 1,
    "FEB": 2, "FEBRUARY": 2,
    "MAR": 3, "MARCH": 3,
    "APR": 4, "APRIL": 4,
    "MAY": 5,
    "JUN": 6, "JUNE": 6,
    "JUL": 7, "JULY": 7,
    "AUG": 8, "AUGUST": 8,
    "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
    "OCT": 10, "OCTOBER": 10,
    "NOV": 11, "NOVEMBER": 11,
    "DEC": 12, "DECEMBER": 12,
}

def detect_month_from_filename(filename: str) -> int:
    name_upper = filename.upper()
    for key, month_num in MONTH_MAP.items():
        if key in name_upper:
            return month_num
    return 99

def find_header_positions(df: pd.DataFrame) -> dict:
    header_map = {"date": None, "desc": None, "amount": None}
    for i, row in df.iterrows():
        row_str = [str(x).upper() if pd.notna(x) else "" for x in row]
        if any("ENTRY DATE" in c for c in row_str) and any("DESCRIPTION" in c for c in row_str):
            for j, cell in enumerate(row_str):
                if "ENTRY DATE" in cell:
                    header_map["date"] = j
                if "DESCRIPTION" in cell:
                    header_map["desc"] = j
                if "AMOUNT" in cell:
                    header_map["amount"] = j
            return header_map, i
    return header_map, 0

def extract_excel_transactions(file_path: str) -> pd.DataFrame:
    raw_df = pd.read_excel(file_path, sheet_name=0, header=None)
    header_map, header_row = find_header_positions(raw_df)
    df = raw_df.iloc[header_row+1:].reset_index(drop=True)

    records = []
    or_counter = 1
    pv_counter = 1
    current_desc = []
    last_entry_date = None

    for _, row in df.iterrows():
        entry_date = row.iloc[header_map["date"]] if header_map["date"] is not None and len(row) > header_map["date"] else None
        description = row.iloc[header_map["desc"]] if header_map["desc"] is not None and len(row) > header_map["desc"] else None
        amount = row.iloc[header_map["amount"]] if header_map["amount"] is not None and len(row) > header_map["amount"] else None

        if pd.notna(entry_date):
            last_entry_date = entry_date

        if isinstance(description, str) and "ENDING BALANCE" in description.upper():
            if current_desc and records:
                merged_desc = " ".join([d for d in current_desc if pd.notna(d)])
                records[-1]["Extracted Description"] += " " + merged_desc
                current_desc = []
            break

        if isinstance(description, str) and "BEGINNING BALANCE" in description.upper():
            continue
        if isinstance(amount, str) and "TRANSACTION AMOUNT" in str(amount).upper():
            continue

        if last_entry_date is not None and pd.notna(amount):
            if current_desc and records:
                merged_desc = " ".join([d for d in current_desc if pd.notna(d)])
                records[-1]["Extracted Description"] += " " + merged_desc
                current_desc = []

            amount_str = str(amount).replace(",", "").strip()
            try:
                if amount_str.endswith("+"):
                    numeric_amount = float(amount_str[:-1])
                elif amount_str.endswith("-"):
                    numeric_amount = -float(amount_str[:-1])
                else:
                    numeric_amount = float(amount_str)
            except ValueError:
                continue

            if numeric_amount >= 0:
                doc_no = f"OR{or_counter}"
                doc_type = "OR"
                or_counter += 1
            else:
                doc_no = f"PV{pv_counter}"
                doc_type = "PV"
                pv_counter += 1

            date_str = str(last_entry_date).strip()
            if re.match(r"^\d{2}/\d{2}$", date_str):
                date_str = f"{date_str}/{YEAR}"
            doc_date = pd.to_datetime(date_str, dayfirst=True, errors="coerce")

            if pd.notna(doc_date):
                try:
                    doc_date = doc_date.strftime("%-d/%-m/%Y")
                except:
                    doc_date = doc_date.strftime("%#d/%#m/%Y")
            else:
                doc_date = ""

            record = {
                "DocNo": doc_no,
                "DocNo2": "",
                "DocDate": doc_date,
                "TaxDate": "",
                "DocType": doc_type,
                "JournalType": "Bank",
                "DealWith": "",
                "TaxEntity": "",
                "Description": "",
                "Extracted Description": str(description).strip() if pd.notna(description) else "",
                "CurrencyCode": "",
                "PaymentAmt": "",
                "BankCharge": "",
                "BankChargeTaxCode": "",
                "BankChargeTaxRate": "",
                "BankChargeTax": "",
                "ToBankRate": "",
                "ToAccountRate": numeric_amount,
                "BankChargeTaxRefNo": "",
                "PaymentBy": "",
                "FloatDay": "",
                "BankChargeDeptNo": ""
            }
            records.append(record)

        elif pd.isna(amount) and pd.notna(description):
            current_desc.append(str(description).strip())

    return pd.DataFrame(records, columns=TEMPLATE_COLUMNS)

def process_excel_files():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    all_records = []
    month_summary = {}
    files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith(".xlsx")]
    files_sorted = sorted(files, key=lambda f: detect_month_from_filename(f))

    for idx, file in enumerate(files_sorted, start=1):
        file_path = os.path.join(DATA_DIR, file)
        month_num = detect_month_from_filename(file)
        month_name = [k for k, v in MONTH_MAP.items() if v == month_num and len(k) > 3]
        month_label = month_name[0].capitalize() if month_name else "Unknown"
        print(f"[{idx}/{len(files_sorted)}] 📂 {file} → Detected month: {month_label}")

        try:
            df = extract_excel_transactions(file_path)
            if not df.empty:
                all_records.append(df)
                month_summary[month_label] = month_summary.get(month_label, 0) + len(df)
            else:
                print(f"⚠️ No valid transactions found in {file}")
        except Exception as e:
            print(f"⚠️ Skipped {file} due to error: {e}")

    if all_records:
        final_df = pd.concat(all_records, ignore_index=True)
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n✅ Processed {len(final_df)} transactions from {len(files_sorted)} files → {OUTPUT_FILE}")

        print("\n📊 Monthly Transaction Summary:")
        for month, count in month_summary.items():
            print(f"   {month}: {count} transactions")
    else:
        print("No valid transactions extracted.")

if __name__ == "__main__":
    process_excel_files()
