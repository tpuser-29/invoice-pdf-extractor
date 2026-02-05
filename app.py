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
                row["Upload_Time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                combined_data.append(row)

        except Exception as e:
            st.error(f"Failed to process {file.name}: {e}")

    if combined_data:
        df = pd.DataFrame(combined_data)

        st.subheader("📊 Merged Invoice Line Items")
        st.dataframe(df, use_container_width=True)

        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)

        st.download_button(
            label="📥 Download Excel (Merged)",
            data=buffer.getvalue(),
            file_name="merged_invoice_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No tabular data found in uploaded PDFs.")
