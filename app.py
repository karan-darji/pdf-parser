import os
import pdfplumber
import pandas as pd
import logging
import json
from datetime import datetime


now = datetime.now()
PDF_DIR = "pdfs"
OUTPUT_EXCEL = f"extracted_data_for_{now.strftime('%Y-%m-%d %H-%M-%S')}.xlsx"


import logging

logging.basicConfig(
    filename='app.log',          # Log file name
    filemode='w',                # 'w' = overwrite, 'a' = append
    level=logging.INFO,         # Minimum log level
    format='%(asctime)s - %(levelname)s - %(message)s'
)


# Fields we want to extract
TARGET_FIELDS = [
    "CSBNumber", "FillingDate", "CourierRegistrationNum-ber", "CourierName", "City",
    "State","Airlines", "PortofLoading", "HAWBNumber", "DeclaredWeight(inKgs)",
    "ImportExportCode(IEC)", "InvoiceTerm", "ExportUsinge-Commerce", "ADCode",
    "Government/Non-Government", "Status", "FOBValue(InINR)",
    "FOBExchangeRate(InFor-eignCurrency)", "EGMNumber", "NameoftheConsignor",
    "NameoftheConsignee", "StateCode", "InvoiceNumber", "InvoiceDate", "CTSH",
    "Quantity", "UnitPrice", "UnitPriceCurrency", "TotalItemValue(InINR)",
    "TaxableValueCurrency", "BONDORUT", "CRNNumber", "Postal/ZipCode",
    "FlightNumber", "DateofDeparture", "NumberofPackages/Pieces/Bags/ULD",
    "MHBSNo", "AirportofDestination", "AccountNo", "NFEI", "LEODATE",
    "EGMDate", "KYCID", "UnitOfMeasure", "ExchangeRate", "CRNMHBSNumber"
]

SECOND_LINE_FIELDS = [
 "InvoiceNumber","InvoiceDate","InvoiceValue(inINR)","CRNNumber","CRNMHBSNumber"
]

'''TARGET_FIELDS = [
    "ExchangeRate"
]'''
SPECIAL_FIELDS = [
    "UnitOfMeasure","ExchangeRate"
]

def match_table_keys(table):
    data = {}
    for row_id,row in enumerate(table):
        if not row or len(row) < 2:
            continue
        if row_id + 1 < len(table):
            next_row = table[row_id + 1]
        else:
            next_row = []
        for idx, val in enumerate(row):
            if not val:
                continue
            label = val.strip().replace(":", "").replace("\n","").strip()
            # logging.info(f"main label: {label}")
            if(label in SECOND_LINE_FIELDS):
                value = ""
                for n_idx, n_val in enumerate(next_row):
                    # logging.info(f"idx: {idx} new idx: {n_idx} val: {n_val} evalute {idx!= n_idx}")
                    if not n_val or idx!= n_idx:
                        continue
                    value = n_val.strip().replace(":", "").replace("\n","").strip()
                    break
            else:
                if label in SPECIAL_FIELDS:
                    value = row[idx + 2].strip() if idx + 2 < len(row) and row[idx + 2] else ""
                else:
                    value = row[idx + 1].strip() if idx + 1 < len(row) and row[idx + 1] else ""
            value = value.replace("\n", "").strip()
            # Try to match based on partial label (case insensitive)

            # logging.info(f"label = {label}, value = {value}")
            for field in TARGET_FIELDS:                
                # logging.info(f"Checking if '{label.lower()}' matches '{field.lower()}' with value '{value}'")    
                if label.lower() == field.lower() and field not in data:
                    # print(f"Checking if '{label.lower()}' in '{field.lower()}'")
                    data[field] = value
    logging.info(f"prepared data: {json.dumps(data)}")
    return data

def extract_from_pdf(path):
    extracted = {}

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            try:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    data = match_table_keys(table)
                    extracted.update(data)
            except:
                continue
    # Fill missing fields as blank
    return {field: extracted.get(field, "") for field in TARGET_FIELDS}

def main():
    all_records = []

    for filename in os.listdir(PDF_DIR):
        if filename.lower().endswith(".pdf"):
            path = os.path.join(PDF_DIR, filename)
            record = extract_from_pdf(path)
            all_records.append(record)
    logging.info(f"Extracted {json.dumps(all_records)}")
    df = pd.DataFrame(all_records)
    # logging.info(f"Final dataframe {df}")
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"Data extraction complete. Written to: {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()
