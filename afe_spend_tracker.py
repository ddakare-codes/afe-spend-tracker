import pdfplumber
import pandas as pd
import os
import time
import shutil
import numpy as np
import re
from datetime import datetime, timezone
import warnings

PDF_FOLDER = r"input_pdfs/"
PROCESSED_FOLDER = os.path.join(PDF_FOLDER, "Processed")
XLSX_PATH = r"data/AFE Spend Tracker.xlsx"

os.makedirs(PROCESSED_FOLDER, exist_ok=True)

COLUMN_MAPPING = {
    "AFE": "HNSAFE",
    "AFE Description": "Project Name (AFE Description)",
    "Start Date": "Start date",
    "End Date": "End date",
    "Investment Reason": "HNS GPB Budget Line from AFE",
    "Total Cost $": "HNS AFE Amount ($K) GROSS",
    "Pdf Type": "AFE PDF TYPE"
}

mapping_dict = {
    "Base O&M (Expense - Facility)": ("OM", 94),
    "Capital - Workover": ("OW", 91),
    "Information Only": ("OM", 94),
    "Kid": ("OM", 94),
    "Major Repairs (Expense - Facility Repairs)": ("OM", 94),
    "Expense - Facility Repair": ("OM", 94),
    "Expense - Facility": ("OM", 94),
    "Wellwork": ("OW", 91),
    "Expense - Workover": ("OW", 91),
    "EH&S": ("OA", 92),
    "P&A": ("OA", 92),
    "Studies": ("OJ", 97),
    "EH&S Work": ("OA", 92)
}

mapping_dict_2 = {
    "Base O&M (Expense - Facility)": ("Major Repairs", "Major Repairs", "Major Repair OPEX"),
    "Capital - Workover": ("Resource Capital - WW", "Wellwork", "Exp-reclassified from CAP"),
    "Major Repairs (Expense - Facility Repairs)": ("Major Repairs", "Major Repairs", "Major Repair OPEX"),
    "Expense - Workover": ("Wellwork", "Wellwork", "Exp"),
    "EH&S": ("Enviro/Decom", "Enviro/Decom", "Exp"),
    "P&A": ("Enviro/Decom", "Enviro/Decom", "Exp"),
    "Studies": ("Studies", "Studies", "ESIB OPEX"),
    "Expense - Facility": ("Major Repairs", "Major Repairs", "Major Repair OPEX"),
    "Expense - Facility Repair": ("Major Repairs", "Major Repairs", "Major Repair OPEX"),
    "EH&S Work": ("Enviro/Decom", "Enviro/Decom", "Exp")
}

def process_page_table(page):
    tables = page.extract_tables()
    structured_data = {}
    page_text = page.extract_text().lower() if page.extract_text() else ""
    pdf_type = "SUP" if "supplement" in page_text else ""

    for table in tables:
        if not table:
            continue
        for i in range(len(table) - 1):
            row = table[i]
            next_row = table[i + 1]
            if not row or not next_row:
                continue
            for j in range(len(row)):
                key = row[j].strip() if row[j] else None
                value = next_row[j].strip() if j < len(next_row) and next_row[j] else None
                if key and value:
                    structured_data[key] = value

    structured_data["Pdf Type"] = pdf_type
    return structured_data

def clean_data(data):
    for key, value in data.items():
        if isinstance(value, str):
            data[key] = re.sub(r'[^\x20-\x7E]', '', value)
    return data

def correct_year(date_str):
    if pd.isnull(date_str):
        return date_str
    try:
        return pd.to_datetime(date_str, errors='coerce')
    except:
        return pd.to_datetime(str(date_str)[:10], errors='coerce')

def process_pdf():
    pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("No PDF found")
        return

    for pdf_file in pdf_files:
        pdf_path = os.path.join(PDF_FOLDER, pdf_file)
        print(f"Processing {pdf_file}")

        try:
            extracted_data = {}

            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    extracted_data.update(process_page_table(page))

            mapped_data = {
                COLUMN_MAPPING[key]: value
                for key, value in extracted_data.items()
                if key in COLUMN_MAPPING
            }

            for key in COLUMN_MAPPING.values():
                if key not in mapped_data:
                    mapped_data[key] = None

            mapped_data = clean_data(mapped_data)
            df_new = pd.DataFrame([mapped_data])

            df_existing = pd.read_excel(XLSX_PATH, sheet_name='PBU Hilcorp')

            df_new["Project ID"] = df_new["HNS GPB Budget Line from AFE"].apply(
                lambda x: mapping_dict.get(x, ("", ""))[0]
            )

            df_new["Project Type"] = df_new["HNS GPB Budget Line from AFE"].apply(
                lambda x: mapping_dict.get(x, ("", ""))[1]
            )

            df_new["PROJECTS- HNS GPB Budget line"] = df_new["HNS GPB Budget Line from AFE"].apply(
                lambda x: mapping_dict_2.get(x, ("", "", ""))[0]
            )

            df_new["HNS-CAP or EXP"] = df_new["HNS GPB Budget Line from AFE"].apply(
                lambda x: "CAP" if isinstance(x, str) and "capital" in x.lower() else "EXP"
            )

            final_df = pd.concat([df_existing, df_new], ignore_index=True)

            final_df["Start date"] = final_df["Start date"].apply(correct_year)
            final_df["End date"] = final_df["End date"].apply(correct_year)

            final_df["Timestamp (UTC)"] = datetime.now(timezone.utc).replace(tzinfo=None)

            with pd.ExcelWriter(XLSX_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                final_df.to_excel(writer, sheet_name="AFE Sheet", index=False)

            shutil.move(pdf_path, os.path.join(PROCESSED_FOLDER, pdf_file))

            print("Done")

        except Exception as e:
            print(f"Error processing {pdf_file}: {e}")

if __name__ == "__main__":
    process_pdf()
