import streamlit as st
import pandas as pd
import os
import io

# App config
st.set_page_config(page_title="Sherawali Agency - Excel Merger", layout="centered")

# Header
st.title("📁 Sherawali Agency - Excel Auto Merger Tool")
st.markdown("Developed by **Ruchi** | Owners: **Santosh Tiwari** & **Krishna Tiwari**")
st.markdown("---")

# Ignore keywords
ignore_keywords = ['merged', 'updated list']

# Column map
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

# File uploader
uploaded_files = st.file_uploader("📤 Upload Excel Files", type=['xls', 'xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🔄 Merge Files"):
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

            # Convert to Excel in memory using xlsxwriter
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False)
            excel_buffer.seek(0)

            st.success("✅ Files merged successfully!")
            st.download_button(
                label="📥 Download Merged Excel File",
                data=excel_buffer,
                file_name="Updated_List.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.markdown("### ✅ Merged Sheets From:")
            for item in merged_files:
                st.markdown(f"- {item}")
        else:
            st.warning("⚠️ No data found to merge from uploaded files.")
else:
    st.info("📂 Please upload Excel files to begin merging.")

# Footer
st.markdown("---")
st.markdown("Developed with ❤️ by **Ruchi** for **Sherawali Agency**")
