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

# Segmented views
def show_segmented_view(full_df):
    view = st.radio("📂 View Mode", ["All", "Buyer", "Seller"], horizontal=True)

    if view == "Buyer":
        df = full_df[full_df["Buyer"].notna()]
    elif view == "Seller":
        df = full_df[full_df["Seller"].notna()]
    else:
        df = full_df

    filtered_df = apply_filters(df, full_df)
    st.markdown(f"### 📋 Showing: {view} View")
    st.dataframe(filtered_df, use_container_width=True)
    export_buttons(filtered_df)
    
def apply_filters(df, full_df):
    st.markdown("### 🔍 Advanced Filters")

    col1, col2 = st.columns(2)
    with col1:
        shape = st.multiselect("Shape", full_df["Shape"].dropna().unique())
        color = st.multiselect("Color", full_df["Color"].dropna().unique())
        quality = st.multiselect("Quality", full_df["Quality"].dropna().unique())
        seller = st.multiselect("Seller", full_df["Seller"].dropna().unique())
    with col2:
        buyer = st.multiselect("Buyer", full_df["Buyer"].dropna().unique())
        pointer = st.multiselect("Pointer", full_df["Pointer"].dropna().unique())
        size = st.multiselect("Size (mm)", full_df["Size (mm)"].dropna().unique())
        date = st.multiselect("Date", full_df["Date"].dropna().unique())

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
    else:
        mask = pd.Series([False] * len(df))
        for col, values in filters.items():
            if values:
                mask |= df[col].isin(values)
        df = df[mask]

    return df
# Main logic
# Main logic
if uploaded_file:
    full_df = load_data(uploaded_file)  # ✅ Add this line here

    if not full_df.empty:
        show_segmented_view(full_df)  # Pass full_df into your view function
else:
    st.info("📎 Please upload an Excel file to begin.")




