import streamlit as st
import pandas as pd
from openpyxl import load_workbook

import streamlit as st
import pandas as pd

st.title("📊 Buyer-Seller Dashboard")

uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xls", "xlsx", "xlsm"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("File uploaded successfully!")
    
    # Use df in your dashboard logic
    st.dataframe(df)  # Replace with your custom views
else:
    st.warning("Please upload an Excel file to proceed.")
SHEET_NAME = "DATA"

def load_data():
from openpyxl import load_workbook
import pandas as pd

def load_data(uploaded_file):
    wb = load_workbook(uploaded_file, read_only=True, keep_vba=True)
    
    if "DATA" not in wb.sheetnames:
        st.error("❌ Sheet named 'DATA' not found in the uploaded file.")
        return pd.DataFrame()  # Return empty DataFrame to avoid crash

    sheet = wb["DATA"]
    data = list(sheet.values)

    # Optional: convert to DataFrame with headers
    headers = data[0]
    rows = data[1:]
    df = pd.DataFrame(rows, columns=headers)
    return df

def advanced_filter(df, column, filter_text):
    if not filter_text:
        return df
    filter_text = filter_text.lower()
    or_groups = [group.strip() for group in filter_text.split("or")]
    def row_matches(value):
        val = str(value).lower()
        for group in or_groups:
            and_terms = [term.strip() for term in group.split("and")]
            if all(term in val for term in and_terms):
                return True
        return False
    return df[df[column].apply(row_matches)]

def main():
    st.set_page_config(page_title="Excel Dashboard", layout="wide")
    st.title("📊 Excel Data Dashboard")

    df = load_data()
    headers = df.columns.tolist()

    st.sidebar.header("🧭 Choose Data View")
    view_choice = st.sidebar.radio("Select view type:", ["Buyer Data", "Seller Data", "A Data"])

    # Apply filters
    st.sidebar.header("🔎 Filter Parameters")
    filters = {
        "Shape": st.sidebar.text_input("Shape"),
        "Size (mm)": st.sidebar.text_input("Size (mm)"),
        "Sieve": st.sidebar.text_input("Sieve"),
        "Pointer": st.sidebar.text_input("Pointer"),
        "Color": st.sidebar.text_input("Color"),
        "Quality": st.sidebar.text_input("Quality"),
        "Seller": st.sidebar.text_input("Seller"),
        "Buyer": st.sidebar.text_input("Buyer"),
        "Date": st.sidebar.text_input("Date")
    }

    start_date = st.sidebar.text_input("Start Date (dd-mm-yyyy)")
    end_date = st.sidebar.text_input("End Date (dd-mm-yyyy)")

    for col, val in filters.items():
        df = advanced_filter(df, col, val)

    # 📅 Apply date range filter
    if start_date and end_date:
        try:
         start_dt = pd.to_datetime(start_date, format="%d-%m-%Y", errors="coerce")
         end_dt = pd.to_datetime(end_date, format="%d-%m-%Y", errors="coerce")
         df["Date"] = pd.to_datetime(df["Date"], errors="coerce")  # ✅ This line is critical
         df = df[(df["Date"] >= start_dt) & (df["Date"] <= end_dt)]
        except:
             st.warning("⚠️ Invalid date format. Use dd-mm-yyyy.")

    # 🗓️ Format Date column
    # 🗓️ Safely format Date column
    if "Date" in df.columns:
         df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
         df["Date"] = df["Date"].dt.strftime("%d-%m-%Y")

# 💯 Correct Terms formatting (Excel already stores it as decimal for %)
    if "Terms" in df.columns:
        try:
            df["Terms"] = pd.to_numeric(df["Terms"], errors="coerce")
            df["Terms"] = df["Terms"].map(lambda x: f"{x * 100:.2f}%" if pd.notnull(x) else "")
        except Exception as e:
            st.warning(f"⚠️ Could not format Terms column: {e}")

# Select columns based on view
    if view_choice == "Buyer Data":
        df = df.iloc[:, :14]
    elif view_choice == "Seller Data":
        df = pd.concat([df.iloc[:, :7], df.iloc[:, 14:20]], axis=1)

    st.subheader("🔍 Filtered Results")
    st.dataframe(df, use_container_width=True)
    st.markdown(f"**Total Entries: {len(df)}**")

    filename = st.text_input("📁 Export filename (without extension)")
    if st.button("📤 Export to Excel"):
        if filename:
            export_path = f"C:/ExcelTest/{filename}.xls"
            df.to_excel(export_path, index=False, engine="openpyxl")
            st.success(f"Exported to {export_path}")
        else:
            st.error("Please enter a filename.")

if __name__ == "__main__":

    main()

