import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime
import io

st.set_page_config(page_title="Invoice Extractor", layout="wide")

st.title("📄 Invoice PDF → Excel Converter")

uploaded_files = st.file_uploader(
    "Drag & drop invoice PDFs",
    type="pdf",
    accept_multiple_files=True
)

def extract_text(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            if page.extract_text():
                text += page.extract_text() + "\n"
    return text

if uploaded_files:
    data = []

    for file in uploaded_files:
        text = extract_text(file)

        data.append({
            "Filename": file.name,
            "Extracted Text": text[:500],  # preview
            "Uploaded At": datetime.now()
        })

    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True)

    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    st.download_button(
        "📥 Download Excel",
        excel_buffer.getvalue(),
        file_name="invoice_data.xlsx"
    )

