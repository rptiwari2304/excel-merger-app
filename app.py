import streamlit as st
import pandas as pd
import zipfile
import io
import os
from datetime import datetime

st.set_page_config(page_title="Sherawali Agency - Excel Merger", layout="centered")
st.title("📂 Sherawali Agency - Excel Auto Merger Tool")
st.markdown("Owner: **Santosh Tiwari** and **Krishna Tiwari**  |  Developer: **ER Ruchi Tiwari**")

ignore_keywords = ['merged', 'updated list']

column_map = {
    'Customer Name': ['cust', 'name', 'customer', 'person', 'people'],
    'Chassis Number': ['chassis', 'cha', 'ch no', 'chsno'],
    'Engine Number': ['engine', 'eng no', 'e no', 'engan'],
    'Registration Number': ['reg no', 'vehicle no', 'rc number', 'registration']
}

def find_best_match(columns, keywords):
    for col in columns:
        col_clean = str(col).lower().strip()
        for keyword in keywords:
            if keyword in col_clean:
                return col
    return None

uploaded_zip = st.file_uploader("📥 Upload a ZIP file containing Excel files", type=['zip'])

if uploaded_zip and st.button("🔄 Merge Files"):
    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        file_list = zip_ref.namelist()
        merged_data = []
        merged_files = []

        for file in file_list:
            if not any(kw in file.lower() for kw in ignore_keywords) and file.endswith(('.xls', '.xlsx')):
                try:
                    with zip_ref.open(file) as f:
                        excel_file = pd.ExcelFile(f)
                        for sheet_name in excel_file.sheet_names:
                            df = excel_file.parse(sheet_name)
                            df_columns = df.columns.tolist()
                            selected_cols = {}

                            for std_col, variations in column_map.items():
                                match = find_best_match(df_columns, variations)
                                selected_cols[std_col] = match

                            new_df = pd.DataFrame()
                            for col in ['Customer Name', 'Chassis Number', 'Engine Number', 'Registration Number']:
                                if selected_cols[col]:
                                    new_df[col] = df[selected_cols[col]]
                                else:
                                    new_df[col] = 'NOT Available'

                            new_df.replace('', 'NA', inplace=True)
                            new_df.fillna('NA', inplace=True)

                            merged_data.append(new_df)
                            merged_files.append(f"{file} → {sheet_name}")
                except Exception as e:
                    st.error(f"❌ Error reading file {file}: {e}")

        if merged_data:
            final_df = pd.concat(merged_data, ignore_index=True)
            final_df['1st Confirmer Name'] = 'Krishna Tiwari'
            final_df['1st Confirmer Mobile Number'] = '9993654016'
            final_df['2nd Confirmer Name'] = 'Santosh Tiwari'
            final_df['2nd Confirmer Mobile Number'] = '9302464234'

            max_rows = 1048575
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for i in range(0, len(final_df), max_rows):
                    part_df = final_df.iloc[i:i+max_rows]
                    excel_io = io.BytesIO()
                    with pd.ExcelWriter(excel_io, engine='xlsxwriter') as writer:
                        part_df.to_excel(writer, index=False)
                    excel_io.seek(0)
                    zip_out.writestr(f"Merged_Part_{i//max_rows + 1}.xlsx", excel_io.read())

            st.success("✅ Files merged and zipped successfully!")
            st.download_button(
                label="📦 Download Merged ZIP File",
                data=zip_buffer.getvalue(),
                file_name="Merged_Excel_Files.zip",
                mime="application/zip"
            )

            st.markdown("### 🔍 Merged from:")
            for item in merged_files:
                st.markdown(f"- {item}")
        else:
            st.warning("⚠️ No valid data found to merge.")
else:
    st.info("📤 Please upload a ZIP file to begin merging.")
