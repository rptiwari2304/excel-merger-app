import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Ignore keywords
ignore_keywords = ['merged', 'updated list']

# Column map as per your code
column_map = {
    'Customer Name': [
        'cust', 'Name', 'person', 'people', 'customer', 'cast', 'cast name', 'castnam',
        'customar', 'kustomer', 'custmr', 'costomer','name','PPL N',
        'cst name', 'custumer', 'customer nam', 'nam of cust', 'nam', 'person name',
        'castnme', 'cstm name', 'castomer', 'nme', 'castnami', 'cstmr name'
    ],
    'Chassis Number': [
        'chassis No', 'cha', 'c no', 'chasis', 'chassis number', 'chasis no', 'chacis',
        'chacis number', 'chassie', 'chas no', 'chas num', 'ch no', 'chas numbr','CHSNO',
        'chass num', 'chasnam', 'cha no', 'cha num', 'che no', 'chessis', 'chasnum',
        'chasy no', 'chasisname', 'chas number', 'chas n', 'chassi', 'chas_n', 'chasn'
    ],
    'Engine Number': [
        'Engine No', 'engin no','ENGNO','engan', 'engan number', 'engan nambar', 'engine num','ENGINE NO',
        'engan no', 'eng no', 'eng num', 'engineno', 'engine no', 'engineno.', 'e no',
        'enjin no', 'engin numbr', 'engineno#', 'enginumber', 'eng no.', 'engn num',
        'e num', 'en num', 'enjin num', 'eng', 'enigne', 'engn'
    ],
    'Registration Number': [
        'reg no', 'regn no', 'registration', 'reg number', 'reg num', 'vehicle reg', 'vehicle reg no',
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
        'register number', 'registered no', 'vehicleregid', 'rcident', 'rc reg no','REGNO'
    ]
}

# Matching function
def find_best_match(columns, keywords):
    for col in columns:
        col_clean = str(col).lower().strip()
        for keyword in keywords:
            if keyword in col_clean:
                return col
    return None

# Streamlit UI
st.title("📂 Excel Merger with Auto Column Matching")
st.markdown("Upload multiple Excel files. Files with 'merged' or 'updated list' in filename will be ignored automatically.")

uploaded_files = st.file_uploader("Upload Excel Files", type=['xls', 'xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("Merge Files"):
        merged_data = []
        merged_files = []

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
                                new_df[col] = 'NOT Available in the list'

                        new_df.replace('', 'NA', inplace=True)
                        new_df.fillna('NA', inplace=True)

                        merged_data.append(new_df)
                        merged_files.append(f"{uploaded_file.name} -> {sheet_name}")
                except Exception as e:
                    st.error(f"Error reading file {uploaded_file.name}: {e}")

        if merged_data:
            final_df = pd.concat(merged_data, ignore_index=True)

            # Add confirmer columns
            final_df['1st Confirmer Name'] = 'Krishna Tiwari'
            final_df['1st Confirmer Mobile Number'] = '11111'
            final_df['2nd Confirmer Name'] = 'Santosh Tiwari'
            final_df['2nd Confirmer Mobile Number'] = '2222'

            # Save output files to Desktop (split if too large)
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            max_rows = 1048575
            file_count = 1
            saved_files = []

            for i in range(0, len(final_df), max_rows):
                part = final_df.iloc[i:i + max_rows]
                output_filename = f"Updated List v{file_count}.xlsx"
                output_path = os.path.join(desktop_path, output_filename)
                part.to_excel(output_path, index=False)
                saved_files.append(output_path)
                file_count += 1

            st.success(f"✅ Files merged and saved successfully to your Desktop!")
            st.markdown("### Saved Files:")
            for fpath in saved_files:
                st.markdown(f"- `{fpath}`")

            st.markdown("### Merged sheets from:")
            for item in merged_files:
                st.markdown(f"- {item}")

        else:
            st.warning("⚠️ No data found to merge from the uploaded files.")
else:
    st.info("Please upload Excel files to begin merging.")
