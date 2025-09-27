# ==============================
# BANK STATEMENT AUTOMATION: EXTRACTION
# ==============================

import os
import pandas as pd

# ==============================
# Step 1: Define Paths
# ==============================
DATA_DIR = "data"
OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "processed_statement.xlsx")

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
    df = pd.read_excel(file_path, sheet_name="Table 1", header=0)

    records = []
    counter = 1
    current_desc = []

    for idx, row in df.iterrows():
        entry_date = row.iloc[0]   # TARIKH MASUK / ENTRY DATE
        description = row.iloc[4]  # TRANSACTION DESCRIPTION
        amount = row.iloc[13]      # TRANSACTION AMOUNT

        # Skip header-like rows
        if isinstance(description, str) and "BEGINNING BALANCE" in description.upper():
            continue
        if isinstance(amount, str) and "TRANSACTION AMOUNT" in amount.upper():
            continue

        # Valid transaction row
        if pd.notna(entry_date) and pd.notna(amount):
            # Merge leftover description lines to previous record
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

            record = {
                "DocNo": f"OR{counter}",
                "DocNo2": "",
                "DocDate": pd.to_datetime(entry_date, errors="coerce").date() if pd.notna(entry_date) else "",
                "TaxDate": "",
                "DocType": "OR",
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
            counter += 1

        # Continuation row (no date/amount but has description text)
        elif pd.isna(entry_date) and pd.isna(amount) and pd.notna(description):
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
