import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from datetime import datetime

st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice PDF → Structured Excel")

uploaded_files = st.file_uploader(
    "Upload invoice PDF files",
    type="pdf",
    accept_multiple_files=True
)

# -------- HEADER EXTRACTION --------
def extract_header_fields(text):
    def find(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    return {
        "INVOICE NO": find(r"Invoice\s*No\.?\s*[:\-]?\s*(\S+)"),
        "INVOICE DATE": find(r"Invoice\s*Date\s*[:\-]?\s*([0-9\/\-]+)"),
        "DUE DATE": find(r"Due\s*Date\s*[:\-]?\s*([0-9\/\-]+)"),
        "BALANCE DUE": find(r"Balance\s*Due\s*[:\-]?\s*([\$0-9,\.]+)"),
        "CUSTOMER NAME": find(r"^(.*?)\n", text)  # top-left first line
    }

# -------- LINE ITEM EXTRACTION --------
REQUIRED_COLUMNS = [
    "TASK", "EMPLOYEE", "SERVICE", "DATE",
    "QTY", "UNIT", "RATE", "SUBTOTAL"
]

def extract_line_items(pdf_file):
    rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                headers = [h.strip().upper() if h else "" for h in table[0]]

                if not set(REQUIRED_COLUMNS).issubset(set(headers)):
                    continue  # skip unrelated tables

                col_index = {h: headers.index(h) for h in REQUIRED_COLUMNS}

                for row in table[1:]:
                    if not any(row):
                        continue

                    item = {col: row[col_index[col]] for col in REQUIRED_COLUMNS}
                    rows.append(item)

    return rows

# -------- MAIN PROCESS --------
if uploaded_files:
    final_rows = []

    for file in uploaded_files:
        try:
            full_text = ""
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    if page.extract_text():
                        full_text += page.extract_text() + "\n"

            header_data = extract_header_fields(full_text)
            line_items = extract_line_items(file)

            for item in line_items:
                final_rows.append({
                    **item,
                    **header_data,
                    "SOURCE FILE": file.name,
                    "UPLOADED AT": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })

        except Exception as e:
            st.error(f"Error processing {file.name}: {e}")

    if final_rows:
        df = pd.DataFrame(final_rows)

        st.subheader("📊 Extracted Invoice Data")
        st.dataframe(df, use_container_width=True)

        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)

        st.download_button(
            "📥 Download Excel",
            buffer.getvalue(),
            file_name="invoice_tabular_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No matching invoice tables found.")
