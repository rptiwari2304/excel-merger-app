import streamlit as st
import pandas as pd
import io
import zipfile
import tempfile
import os

# App Configuration
st.set_page_config(page_title="Sherawali Agency - Excel Merger", layout="centered")
st.title("📂 Sherawali Agency - Excel Auto Merger Tool")
st.markdown("Owner: **Santosh Tiwari** and **Krishna Tiwari**  |  Developer: **ER Ruchi Tiwari**")

# Ignore Keywords
ignore_keywords = ['merged', 'updated list']

# Column mapping
column_map = {
    'Customer Name': ['cust', 'name', 'customer', 'person', 'people'],
    'Chassis Number': ['chassis', 'cha', 'ch no', 'chsno'],
    'Engine Number': ['engine', 'eng no', 'e no', 'engan'],
    'Registration Number': ['reg no', 'vehicle no', 'rc number', 'registration']
}

# Column match helper
def find_best_match(columns, keywords):
    for col in columns:
        col_clean = str(col).lower().strip()
        for keyword in keywords:
            if keyword in col_clean:
                return col
    return None

# Upload zip
uploaded_zip = st.file_uploader("📁 Upload ZIP file containing Excel files", type=["zip"])

if uploaded_zip:
    if st.button("🔄 Merge Excel Files from ZIP"):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)

            merged_data = []
            merged_files = []

            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    if file.endswith((".xls", ".xlsx")) and not any(kw in file.lower() for kw in ignore_keywords):
                        file_path = os.path.join(root, file)
                        try:
                            excel_file = pd.ExcelFile(file_path)
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
                            st.error(f"❌ Error reading {file}: {e}")

            if merged_data:
                final_df = pd.concat(merged_data, ignore_index=True)

                # Add Confirmers
                final_df['1st Confirmer Name'] = 'Krishna Tiwari'
                final_df['1st Confirmer Mobile Number'] = '9993654016'
                final_df['2nd Confirmer Name'] = 'Santosh Tiwari'
                final_df['2nd Confirmer Mobile Number'] = '9302464234'

                # Split by Excel row limit
                max_rows = 1048575
                buffer_list = []

                for i in range(0, len(final_df), max_rows):
                    part_df = final_df.iloc[i:i+max_rows]
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
                st.warning("⚠️ No valid Excel files found in ZIP.")
