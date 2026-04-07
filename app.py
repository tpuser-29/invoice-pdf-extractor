import streamlit as st
import pandas as pd
 
st.set_page_config(page_title="Smart Data Merger ULTIMATE", layout="wide")
 
st.title("🚀 Smart Excel Group Merger ULTIMATE FINAL")
 
uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)
 
remove_duplicates = st.checkbox("Remove Duplicates (Based on Order ID)", value=False)
 
# -------- HEADER DETECTION --------
def detect_header(df):
    for i in range(min(10, len(df))):
        if df.iloc[i].notna().sum() > 2:
            return i
    return 0
 
# -------- CLEAN COLUMN NAMES --------
def clean_columns(cols):
    return [str(c).strip().lower().replace(" ", "").replace("_", "") for c in cols]
 
# -------- STANDARD COLUMN MAPPING --------
def standardize_columns(cols):
    mapping = {}
    for col in cols:
        if "order" in col and "id" in col:
            mapping[col] = "order_id"
        elif "awb" in col:
            mapping[col] = "awb"
        elif "date" in col:
            mapping[col] = "date"
        elif "client" in col or "merchant" in col:
            mapping[col] = "client"
        else:
            mapping[col] = col
    return mapping
 
# -------- MONTH MAP --------
month_map = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
    "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
    "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
}
 
# -------- EXTRACT GROUP + MONTH --------
def extract_info(sheet_name):
    parts = sheet_name.split("-")
    if len(parts) >= 2:
        month = parts[0]
        group = parts[1].lower()
        month_name = month_map.get(month, month)
        return group, month_name
    return sheet_name.lower(), "Unknown"
 
# -------- MAIN --------
if uploaded_files:
 
    grouped_data = {}
 
    for file in uploaded_files:
        xls = pd.ExcelFile(file)
 
        for sheet in xls.sheet_names:
            try:
                raw_df = xls.parse(sheet, header=None)
                header_row = detect_header(raw_df)
 
                df = xls.parse(sheet, header=header_row)
                df.columns = clean_columns(df.columns)
 
                # standardize columns
                mapping = standardize_columns(df.columns)
                df.rename(columns=mapping, inplace=True)
 
                # -------- CLEANING --------
 
                # remove fully empty rows
                df = df.dropna(how="all")
 
                # ✅ REMOVE ROWS WHERE FIRST COLUMN (A) IS BLANK
                first_col = df.columns[0]
                df = df[df[first_col].notna() & (df[first_col].astype(str).str.strip() != "")]
 
                # ✅ REMOVE CLIENT = Commission / Income
                if "client" in df.columns:
                    df = df[~df["client"].astype(str).str.lower().str.contains("commission|income", na=False)]
 
                # -------- GROUP + MONTH --------
                group, month = extract_info(sheet)
                df["month"] = month
 
                if group not in grouped_data:
                    grouped_data[group] = []
 
                grouped_data[group].append(df)
 
            except Exception as e:
                st.warning(f"⚠️ Error in {file.name} - {sheet}")
 
    st.success("✅ Files Processed Successfully")
 
    # -------- PROCESS EACH GROUP --------
    for group, df_list in grouped_data.items():
 
        # ✅ COLUMN ORDER FIX (BASED ON FIRST FILE)
        base_cols = list(df_list[0].columns)
        all_cols = base_cols.copy()
 
        for df in df_list:
            for col in df.columns:
                if col not in all_cols:
                    all_cols.append(col)
 
        final_list = []
        for df in df_list:
            for col in all_cols:
                if col not in df.columns:
                    df[col] = None
            df = df[all_cols]
            final_list.append(df)
 
        final_df = pd.concat(final_list, ignore_index=True)
 
        # -------- DUPLICATE REMOVAL --------
        if remove_duplicates and "order_id" in final_df.columns:
            final_df = final_df.drop_duplicates(subset=["order_id"])
 
        # -------- UI --------
        st.markdown(f"## 📁 {group.upper()}")
 
        col1, col2 = st.columns(2)
        col1.metric("Total Rows", len(final_df))
        col2.metric("Columns", len(final_df.columns))
 
        st.dataframe(final_df.head(100), use_container_width=True)
 
        # -------- DOWNLOAD --------
        csv = final_df.to_csv(index=False).encode("utf-8")
 
        st.download_button(
            f"⬇️ Download {group.upper()}",
            csv,
            f"{group}_merged.csv",
            "text/csv"
        )
 
