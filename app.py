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
    'Customer Name': ['cust', 'name', 'customer', 'person', 'people cust', 'Name', 'person', 'people', 'customer', 'cast', 'cast name', 'castnam',
        'customar', 'kustomer', 'custmr', 'costomer','name','PPL N',
        'cst name', 'custumer', 'customer nam', 'nam of cust', 'nam', 'person name',
        'castnme', 'cstm name', 'castomer', 'nme', 'castnami', 'cstmr name'],
    'Chassis Number': ['chassis', 'cha', 'ch no', 'chsno','chassis No', 'cha', 'c no', 'chasis', 'chassis number', 'chasis no', 'chacis',
        'chacis number', 'chassie', 'chas no', 'chas num', 'ch no', 'chas numbr','CHSNO',
        'chass num', 'chasnam', 'cha no', 'cha num', 'che no', 'chessis', 'chasnum',
        'chasy no', 'chasisname', 'chas number', 'chas n', 'chassi', 'chas_n', 'chasn'],
    'Engine Number': ['engine', 'eng no', 'e no', 'engan','Engine No', 'engin no','ENGNO','engan', 'engan number', 'engan nambar', 'engine num','ENGINE NO',
        'engan no', 'eng no', 'eng num', 'engineno', 'engine no', 'engineno.', 'e no',
        'enjin no', 'engin numbr', 'engineno#', 'enginumber', 'eng no.', 'engn num',
        'e num', 'en num', 'enjin num', 'eng', 'enigne', 'engn'],
    'Registration Number': ['reg no', 'vehicle no', 'rc number', 'registration', 'reg no', 'regn no', 'registration', 'reg number', 'reg num', 'vehicle reg', 'vehicle reg no',
        'regn number', 'reg numb', 'veh reg', 'vehicleno', 'vrn', 'regn', 'regnum', 'veh no',
        'vehicle no', 'vehic no', 'vehregno', 'rc no', 'rc number', 'rcnum', 'rcno', 'registration#',
        'reg#', 'reg. no', 'regno.', 'veh_reg', 'vehicle_num', 'regn.', 'veh no.', 'veh number',
        'rego', 'regstrtn no', 'register no', 'registrationnumb', 'vehicle registration',
        'vehiclereg', 'vehicleregnum', 'vregnum', 'vregno', 'vehicleregno', 'vehicle_reg_no',
        'vehicleregistrationno', 'registrationnumber', 'reg_num', 'v_num', 'vehicle_regnumber',
        'rcnumber', 'rcnumb', 'vehicleregnumbr', 'vehiclereg#', 'regn#', 'vehicleid', 'vehid',
        'vehiclid', 'regid', 'vehicleregno.', 'regnnum', 'reg no.', 'rc_no', 'vehiclenumber',
        'regn numb', 'rc number', 'veh_reg_no', 'vehiclereg no', 'rcno.', 'vehicle registration no',
        'vehicle_registration_num', 'veh_regn', 'v_reg_no', 'vehregnumber', 'rc_num', 'rc_no.',
        'vreg', 'vehiclenum', 'rcid', 'vrno', 'vnum', 'vnumber', 'vnumbr', 'vehcl no', 'vehicl_reg_no',
        'register number', 'registered no', 'vehicleregid', 'rcident', 'rc reg no','REGNO']
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
            final_df['2nd Confirmer Mobile Number'] = '9302465234'

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
