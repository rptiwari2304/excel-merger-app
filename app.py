import streamlit as st
import pandas as pd
import os
import io
import zipfile

# App config
st.set_page_config(page_title="Sherawali Agency - Excel Merger", layout="centered")

# Header
st.title("📁 Sherawali Agency - Excel Auto Merger Tool")
st.markdown("Developed by **Ruchi** | Owners: **Santosh Tiwari** & **Krishna Tiwari**")
st.markdown("---")

ignore_keywords = ['merged', 'updated list']

column_map = {
    'Customer Name': ['cust', 'name', 'customer', 'person'],
    'Chassis Number': ['chassis', 'cha', 'c no'],
    'Engine Number': ['engine no', 'eng', 'e no'],
    'Registration Number': ['reg', 'registration', 'rc']
}

def find_best_match(columns, keywords):
    for col in columns:
        col_clean = str(col).lower().strip()
        for keyword in keywords:
            if keyword in col_clean:
                return col
    return None

uploaded_files = st.file_uploader("📤 Upload Excel Files", type=['xls', 'xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🔄 Merge Files"):
        merged_data = []
        merged_files_info = []

        for uploaded_file in uploaded_files:
            filename = uploaded_file.name.lower()
            if not any(kw in filename for kw in ignore_keywords):
                try:
                    excel_file = pd.ExcelFile(uploaded_file)
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
                        merged_files_info.append(f"{uploaded_file.name} -> {sheet_name}")
                except Exception as e:
                    st.error(f"Error reading {uploaded_file.name}: {e}")

        if merged_data:
            final_df = pd.concat(merged_data, ignore_index=True)

            # Add confirmer columns
            final_df['1st Confirmer Name'] = 'Krishna Tiwari'
            final_df['1st Confirmer Mobile Number'] = '11111'
            final_df['2nd Confirmer Name'] = 'Santosh Tiwari'
            final_df['2nd Confirmer Mobile Number'] = '2222'

            MAX_ROWS = 1048575
            num_parts = (len(final_df) // MAX_ROWS) + 1

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for i in range(num_parts):
                    part_df = final_df[i * MAX_ROWS:(i + 1) * MAX_ROWS]
                    part_buffer = io.BytesIO()
                    with pd.ExcelWriter(part_buffer, engine='openpyxl') as writer:
                        part_df.to_excel(writer, index=False)
                    part_buffer.seek(0)
                    zip_file.writestr(f"Merged_{i+1}.xlsx", part_buffer.read())

            zip_buffer.seek(0)

            st.success("✅ Merging Complete! Download All Files Below.")
            st.download_button(
                label="📦 Download All Merged Excel Files (ZIP)",
                data=zip_buffer,
                file_name="Sherawali_Agency_Merged_Files.zip",
                mime="application/zip"
            )

            st.markdown("### ✅ Merged Sheets From:")
            for item in merged_files_info:
                st.markdown(f"- {item}")
        else:
            st.warning("⚠️ No data found to merge from uploaded files.")
else:
    st.info("📂 Please upload Excel files to begin merging.")

st.markdown("---")
st.markdown("Developed with ❤️ by **Ruchi** for **Sherawali Agency**")
