import polars as pl
import streamlit as st
import msoffcrypto
import json
from datetime import datetime
from io import BytesIO
from resources.excel_tools import save_xlsx, cast_columns

if "merged_accounts" not in st.session_state:
    st.session_state.merged_accounts = None

st.header("Merge Accounts")
st.write(f''':green[:green-background[Masterfile]]''')

with open('./resources/agent_code_bcrm_volare.json', 'r') as file:
    agent_code_bcrm_volare = json.load(file)

with open('./resources/agent_code_volare_fullname.json', 'r') as file:
    agent_code_volare_fullname = json.load(file)

with st.form(key="blasting"):
    file = st.file_uploader("Upload Password Protected XLSX", type=["xlsx"])
    passkey = st.text_input("Password", type="password")

    col1, col2 = st.columns(2, vertical_alignment="bottom")
    col1_1, col2_1 = col1.columns(2, vertical_alignment="bottom")

    # Get the current date
    now = datetime.now()

    # Set the start date to the first of the current month
    start_date = col1_1.date_input("Start Date", value=datetime(now.year, now.month, 1))

    # Set the end date to the current date and time
    end_date = col2_1.date_input("End Date", value="today")

    if col2.form_submit_button("SUBMIT", use_container_width=True):
        if file is not None:
            excel_decrypted = BytesIO()
            with BytesIO(file.read()) as f:
                excel_file = msoffcrypto.OfficeFile(f)
                excel_file.load_key(password = passkey)
                excel_file.decrypt(excel_decrypted)
    
        active_sheet = pl.read_excel(excel_decrypted, sheet_name="ACTIVE", schema_overrides={"ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE": pl.Utf8})
        pout_sheet = pl.read_excel(excel_decrypted, sheet_name="POUT", schema_overrides={"ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE": pl.Utf8})
        #excluded = pl.read_excel(excel_decrypted, sheet_name="EXCLUDED")

        active_sheet = active_sheet.with_columns(
            pl.lit("ACTIVE").alias("REMARKS"),
            pl.lit(None).alias("PULLED OUT DATE")

        )

        active_sheet = active_sheet.select(["PULLED OUT DATE", "REMARKS", "PLACEMENT", "ACCOUNT NUMBER", "ENDO STAT", "CHCODE", "TAGGING", "DPD BUCKET", "DPD_", "MOBILE PROPER", "OB", "FRESH/SPILLOVER", "RECEIVED DATE", "AS_OF", "CPM_ID", "NAME", "FIRST_NAME", "LAST_NAME", "BIRTH_DATE", "ACCOUNT_ID", "PRODUCT_NAME"])

        pout_sheet = pout_sheet.filter((pl.col("PULLED OUT DATE") >= start_date) & (pl.col("PULLED OUT DATE") <= end_date))
        pout_sheet = pout_sheet.select(["PULLED OUT DATE", "REMARKS", "PLACEMENT", "ACCOUNT NUMBER", "ENDO STAT", "CHCODE", "TAGGING", "DPD BUCKET", "DPD_", "MOBILE PROPER", "OB", "FRESH/SPILLOVER", "RECEIVED DATE", "AS_OF", "CPM_ID", "NAME", "FIRST_NAME", "LAST_NAME", "BIRTH_DATE", "ACCOUNT_ID", "PRODUCT_NAME"])
        
        #excluded = excluded.select(["REMARKS", "PLACEMENT", "ACCOUNT NUMBER", "ENDO STAT", "CHCODE", "TAGGING", "DPD BUCKET", "DPD_", "MOBILE PROPER", "OB", "FRESH/SPILLOVER", "RECEIVED DATE", "AS_OF", "CPM_ID", "NAME", "FIRST_NAME", "LAST_NAME", "BIRTH_DATE", "ACCOUNT_ID", "PRODUCT_NAME"])

        column_types = {
        "PULLED OUT DATE": pl.Date,
        "REMARKS": pl.Utf8,
        "PLACEMENT": pl.Utf8,
        "ACCOUNT NUMBER": pl.Int64,
        "ENDO STAT": pl.Utf8,
        "CHCODE": pl.Utf8,
        "TAGGING": pl.Utf8,
        "DPD BUCKET": pl.Utf8,
        "DPD_": pl.Int64,
        "MOBILE PROPER": pl.Utf8,
        "OB": pl.Float64,
        "FRESH/SPILLOVER": pl.Utf8,
        "RECEIVED DATE": pl.Date,
        "AS_OF": pl.Date,
        "CPM_ID": pl.Utf8,
        "NAME": pl.Utf8,
        "FIRST_NAME": pl.Utf8,
        "LAST_NAME": pl.Utf8,
        "BIRTH_DATE": pl.Date,
        "ACCOUNT_ID": pl.Utf8,
        "PRODUCT_NAME": pl.Utf8
        }

        active_sheet = cast_columns(active_sheet, column_types)
        pout_sheet = cast_columns(pout_sheet, column_types)
        #excluded = cast_columns(excluded, column_types)

        merged_accounts = active_sheet.vstack(pout_sheet)
        #merged_accounts = merged_accounts.vstack(excluded)

        merged_accounts = merged_accounts.with_columns(
            pl.col("TAGGING").map_elements(lambda x: agent_code_bcrm_volare.get(x, None), return_dtype=pl.Utf8).alias("VOLARE TAGGING")
        )

        merged_accounts = merged_accounts.with_columns(
            pl.col("VOLARE TAGGING").map_elements(lambda x: agent_code_volare_fullname.get(x, None), return_dtype=pl.Utf8).alias("AGENT NAME")
        )

        merged_accounts = merged_accounts.select(["PULLED OUT DATE", "REMARKS", "PLACEMENT", "ACCOUNT NUMBER", "ENDO STAT", "CHCODE", "TAGGING", "VOLARE TAGGING", "AGENT NAME", "DPD BUCKET", "DPD_", "MOBILE PROPER", "OB", "FRESH/SPILLOVER", "RECEIVED DATE", "AS_OF", "CPM_ID", "NAME", "FIRST_NAME", "LAST_NAME", "BIRTH_DATE", "ACCOUNT_ID", "PRODUCT_NAME"])

        merged_accounts = merged_accounts.filter(pl.col("AS_OF").is_not_null())
        merged_accounts = merged_accounts.sort("AS_OF", descending=True)

        st.session_state.merged_accounts = merged_accounts

if st.session_state.merged_accounts is not None:
    st.dataframe(st.session_state.merged_accounts)

    date = datetime.now().strftime("%m%d%y")

    formatting = {
        "ACCOUNT NUMBER": "0",
        "OB": "0.00",
        "RECEIVED DATE": "mm/dd/yyyy",
        "AS_OF": "mm/dd/yyyy",
        "BIRTH_DATE": "mm/dd/yyyy",
        "PULLED OUT DATE": "mm/dd/yyyy"
    }

    st.download_button(
        label = "Download Merged Accounts",
        data = save_xlsx(st.session_state.merged_accounts, formatting),
        file_name = f"maya_merged_accounts_{date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )