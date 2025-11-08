import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Buyer-Seller Dashboard", layout="wide")
st.title("📊 Buyer-Seller Dashboard")

# Upload Excel file
uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xls", "xlsx", "xlsm"])

# Load only the 'DATA' sheet
def load_data(uploaded_file):
    try:
        wb = load_workbook(uploaded_file, read_only=True, keep_vba=True)
        if "DATA" not in wb.sheetnames:
            st.error("❌ Sheet named 'DATA' not found.")
            return pd.DataFrame()
        sheet = wb["DATA"]
        data = list(sheet.values)
        headers = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=headers)
        return df
    except Exception as e:
        st.error(f"❌ Error loading file: {e}")
        return pd.DataFrame()

# Export buttons
def export_buttons(df):
    st.markdown("### 📤 Export Options")

    default_name = "buyer_seller"
    file_name = st.text_input("📝 Enter export file name (without extension)", value=default_name)

    col1, col2 = st.columns(2)
    with col1:
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Export CSV",
            data=csv,
            file_name=f"{file_name}.csv",
            mime="text/csv"
        )
    with col2:
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='DATA')
        st.download_button(
            label="📥 Export Excel",
            data=excel_buffer.getvalue(),
            file_name=f"{file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Advanced filters
def apply_filters(df):
    st.markdown("### 🔍 Advanced Filters")

    col1, col2 = st.columns(2)
    with col1:
        shape = st.multiselect("Shape", df["Shape"].dropna().unique())
        color = st.multiselect("Color", df["Color"].dropna().unique())
        quality = st.multiselect("Quality", df["Quality"].dropna().unique())
        seller = st.multiselect("Seller", df["Seller"].dropna().unique())
    with col2:
        buyer = st.multiselect("Buyer", df["Buyer"].dropna().unique())
        pointer = st.multiselect("Pointer", df["Pointer"].dropna().unique())
        size = st.multiselect("Size (mm)", df["Size (mm)"].dropna().unique())
        date = st.multiselect("Date", df["Date"].dropna().unique())

    logic = st.radio("Filter Logic", ["AND", "OR"], horizontal=True)

    filters = {
        "Shape": shape,
        "Color": color,
        "Quality": quality,
        "Seller": seller,
        "Buyer": buyer,
        "Pointer": pointer,
        "Size (mm)": size,
        "Date": date
    }

    if logic == "AND":
        for col, values in filters.items():
            if values:
                df = df[df[col].isin(values)]
    else:  # OR logic
        mask = pd.Series([False] * len(df))
        for col, values in filters.items():
            if values:
                mask |= df[col].isin(values)
        df = df[mask]

    return df

# Segmented views
def show_segmented_view(df):
    view = st.radio("📂 View Mode", ["All", "Buyer", "Seller"], horizontal=True)
    if view == "Buyer":
        buyers = df["Buyer"].dropna().unique()
        selected = st.selectbox("Select Buyer", buyers)
        df = df[df["Buyer"] == selected]
    elif view == "Seller":
        sellers = df["Seller"].dropna().unique()
        selected = st.selectbox("Select Seller", sellers)
        df = df[df["Seller"] == selected]

    filtered_df = apply_filters(df)
    st.markdown(f"### 📋 Showing: {view} View")
    st.dataframe(filtered_df, use_container_width=True)
    export_buttons(filtered_df)

# Main logic
if uploaded_file:
    df = load_data(uploaded_file)
    if not df.empty:
        show_segmented_view(df)
else:
    st.info("📎 Please upload an Excel file to begin.")



