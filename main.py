import os
import pandas as pd
import re

# ==============================
# Step 1: Define Paths & Config
# ==============================
DATA_DIR = "data"
OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "processed_statement.xlsx")

# >>> Edit this for the year of the statement <<<
YEAR = "2025"

# ==============================
# Step 2: Define Template Columns
# ==============================
TEMPLATE_COLUMNS = [
    "DocNo", "DocNo2", "DocDate", "TaxDate", "DocType", "JournalType",
    "DealWith", "TaxEntity", "Description", "Extracted Description",
    "CurrencyCode", "PaymentAmt", "BankCharge", "BankChargeTaxCode",
    "BankChargeTaxRate", "BankChargeTax", "ToBankRate", "ToAccountRate",
    "BankChargeTaxRefNo", "PaymentBy", "FloatDay", "BankChargeDeptNo"
]

# ==============================
# Step 3: Extract Transactions from Excel
# ==============================
def extract_excel_transactions(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name=0, header=0)

    records = []
    or_counter = 1
    pv_counter = 1
    current_desc = []
    last_entry_date = None

    for idx, row in df.iterrows():
        entry_date = row.iloc[0]   # ENTRY DATE
        description = row.iloc[4]  # TRANSACTION DESCRIPTION
        amount = row.iloc[13]      # TRANSACTION AMOUNT

        # Track last non-empty date
        if pd.notna(entry_date):
            last_entry_date = entry_date

        # Skip header-like rows
        if isinstance(description, str) and "BEGINNING BALANCE" in description.upper():
            continue
        if isinstance(amount, str) and "TRANSACTION AMOUNT" in amount.upper():
            continue

        # Valid transaction row
        if last_entry_date is not None and pd.notna(amount):
            # Merge leftover description lines into last record
            if current_desc and records:
                merged_desc = " ".join([d for d in current_desc if pd.notna(d)])
                records[-1]["Extracted Description"] += " " + merged_desc
                current_desc = []

            # Clean numeric amount
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

            # Assign DocNo and DocType
            if numeric_amount >= 0:
                doc_no = f"OR{or_counter}"
                doc_type = "OR"
                or_counter += 1
            else:
                doc_no = f"PV{pv_counter}"
                doc_type = "PV"
                pv_counter += 1

            # Build DocDate string with configured YEAR
            date_str = str(last_entry_date).strip()
            if re.match(r"^\d{2}/\d{2}$", date_str):  # DD/MM
                date_str = f"{date_str}/{YEAR}"
            doc_date = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
            doc_date = doc_date.strftime("%d/%m/%Y") if pd.notna(doc_date) else ""

            record = {
                "DocNo": doc_no,
                "DocNo2": "",
                "DocDate": doc_date,
                "TaxDate": "",
                "DocType": doc_type,
                "JournalType": "",
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

        # Continuation row (no amount but description exists)
        elif pd.isna(amount) and pd.notna(description):
            current_desc.append(str(description).strip())

    # Flush last continuation description
    if current_desc and records:
        merged_desc = " ".join([d for d in current_desc if pd.notna(d)])
        records[-1]["Extracted Description"] += " " + merged_desc

    return pd.DataFrame(records, columns=TEMPLATE_COLUMNS)

# ==============================
# Step 4: Process All Excel Files
# ==============================
def process_excel_files():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    all_records = []
    for file in os.listdir(DATA_DIR):
        if file.lower().endswith(".xlsx"):
            file_path = os.path.join(DATA_DIR, file)
            df = extract_excel_transactions(file_path)
            if not df.empty:
                all_records.append(df)

    if all_records:
        final_df = pd.concat(all_records, ignore_index=True)
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"Processed {len(final_df)} transactions → {OUTPUT_FILE}")
    else:
        print("No valid transactions extracted.")

# ==============================
# Run Script
# ==============================
if __name__ == "__main__":
    process_excel_files()
