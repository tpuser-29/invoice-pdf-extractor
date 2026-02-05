import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from datetime import datetime

st.set_page_config(page_title="Invoice PDF → Structured Excel", layout="wide")
st.title("📄 Invoice PDF → Structured Excel")

uploaded_files = st.file_uploader(
    "Upload invoice PDF files",
    type="pdf",
    accept_multiple_files=True
)

# ---------------- HEADER EXTRACTION ----------------
def extract_header_fields(text):

    def find(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    return {
        "INVOICE NO": find(r"Invoice\s*No\.?\s*[:\-]?\s*(\S+)"),
        "INVOICE DATE": find(r"Invoice\s*Date\s*[:\-]?\s*([0-9\/\-]+)"),
        "DUE DATE": find(r"Due\s*Date\s*[:\-]?\s*([0-9\/\-]+)"),
        "BALANCE DUE": find(r"Balance\s*Due\s*[:\-]?\s*([\$0-9,\.]+)"),
        "CUSTOMER NAME": lines[0] if lines else ""
    }

# ---------------- LINE ITEM EXTRACTION ----------------
def extract_line_items_from_text(text):
    rows = []
    lines = [l for l in text.split("\n") if l.strip()]

    start_idx = None

    for i, line in enumerate(lines):
        if "TASK" in line.upper() and "QTY" in line.upper() and "RATE" in line.upper():
            start_idx = i + 1
            break

    if start_idx is None:
        return rows

    for line in lines[start_idx:]:
        if "TOTAL" in line.upper() or "BALANCE" in line.upper():
            break

        parts = re.split(r"\s{2,}", line.strip())

        if len(parts) < 8:
            continue

        rows.append({
            "TASK": parts[0],
            "EMPLOYEE": parts[1],
            "SERVICE": parts[2],
            "DATE": parts[3],
            "QTY": parts[4],
            "UNIT": parts[5],
            "RATE": parts[6],
            "SUBTOTAL": parts[7],
        })

    return rows

# ---------------- MAIN PROCESS ----------------
if uploaded_files:

    final_data = []

    for file in uploaded_files:

        full_text = ""

        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                if page.extract_text():
                    full_text += page.extract_text() + "\n"

        header = extract_header_fields(full_text)
        line_items = extract_line_items_from_text(full_text)

        if not line_items:
            st.warning(f"No line items found in {file.name}")
            continue

        for item in line_items:
            final_data.append({
                **item,
                **header,
                "SOURCE FILE": file.name,
                "UPLOADED AT": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

    if final_data:
        df = pd.DataFrame(final_data)

        st.subheader("📊 Extracted Invoice Line Items")
        st.dataframe(df, use_container_width=True)

        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)

        st.download_button(
            "📥 Download Excel",
            buffer.getvalue(),
            file_name="invoice_structured_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No invoice data extracted.")
