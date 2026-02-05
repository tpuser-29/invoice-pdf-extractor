import streamlit as st
import pandas as pd
import pdfplumber
import io
from datetime import datetime

st.set_page_config(page_title="Invoice PDF Extractor", layout="wide")
st.title("📄 Invoice PDF → Tabular Excel (Merged)")

uploaded_files = st.file_uploader(
    "Upload or drag & drop invoice PDF files",
    type="pdf",
    accept_multiple_files=True
)

def extract_invoice_tables(uploaded_file):
    all_rows = []

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                # First row usually header
                headers = table[0]
                data_rows = table[1:]

                for row in data_rows:
                    if any(row):  # skip empty rows
                        row_dict = dict(zip(headers, row))
                        all_rows.append(row_dict)

    return all_rows

if uploaded_files:
    combined_data = []

    for file in uploaded_files:
        try:
            table_rows = extract_invoice_tables(file)

            for row in table_rows:
                row["Source_File_Name"] = file.name
                row["Upload_Time"] = datetime.now().strftime("%Y-%m-%d_]()
