import streamlit as st
import pandas as pd
import zipfile
import io
import os

# ---------------------- Streamlit Config ----------------------
st.set_page_config(page_title="Sherawali Agency - Excel Merger", layout="centered")
st.title("📂 Sherawali Agency - Excel Auto Merger Tool")
st.markdown("Owner: **Santosh Tiwari** and **Krishna Tiwari**  |  Developer: **ER Ruchi Tiwari**")

# ---------------------- Column Matching ----------------------
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

# ---------------------- Upload ZIP File ----------------------
zip_file = st.file_uploader("📥 Upload a ZIP file containing Excel files", type="zip")

if zip_file:
    if st.button("🔄 Merge Files"):
        merged_data = []
        merged_files = []

        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            file_list = [f for f in zip_ref.namelist() if f.endswith(('.xls', '.xlsx')) and not any(kw in f.lower() for kw in ['merged', 'updated list'])]

            for file in file_list:
                with zip_ref.open(file) as f:
                    try:
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

            # Add confirmer columns
            final_df['1st Confirmer Name'] = 'Krishna Tiwari'
            final_df['1st Confirmer Mobile Number'] = '9993654016'
            final_df['2nd Confirmer Name'] = 'Santosh Tiwari'
            final_df['2nd Confirmer Mobile Number'] = '9302464234'

            max_rows = 1048575
            buffer_list = []

            for i in range(0, len(final_df), max_rows):
                part_df = final_df.iloc[i:i + max_rows]
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    part_df.to_excel(writer, index=False)
                buffer_list.append(excel_buffer)

            st.success("✅ Files merged successfully!")
            for idx, buffer in enumerate(buffer_list, 1):
                st.download_button(
                    label=f"📥 Download Part {idx}",
                    data=buffer.getvalue(),
                    file_name=f"Updated_List_Part_{idx}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.markdown("### 🔍 Merged from:")
            for item in merged_files:
                st.markdown(f"- {item}")
        else:
            st.warning("⚠️ No valid Excel files found in the ZIP.")
else:
    st.info("📤 Please upload a ZIP file to begin merging.")
