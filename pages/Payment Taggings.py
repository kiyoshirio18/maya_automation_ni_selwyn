import polars as pl
import streamlit as st
import os
from io import BytesIO
from datetime import datetime, timedelta

st.header("Payment Taggings")

daily_remark_schema = {
    "S.No": pl.Int64,
    "Date": pl.Date,
    "Time": pl.Datetime,
    "Debtor": pl.Utf8,
    "Account No.": pl.Utf8,
    "Card No.": pl.Utf8,
    "Service No.": pl.Utf8,
    "DPD": pl.Int64,
    "Reason For Default": pl.Utf8,
    "Call Status": pl.Utf8,
    "Status": pl.Utf8,
    "Remark": pl.Utf8,
    "Remark By": pl.Utf8,
    "Remark Type": pl.Utf8,
    "Field Visit Date": pl.Utf8,
    "Collector": pl.Utf8,
    "Client": pl.Utf8,
    "Product Description": pl.Utf8,
    "Product Type": pl.Utf8,
    "Batch No": pl.Utf8,
    "Account Type": pl.Utf8,
    "Relation": pl.Utf8,
    "PTP Amount": pl.Float64,
    "Next Call": pl.Utf8,
    "PTP Date": pl.Utf8,
    "Claim Paid Amount": pl.Float64,
    "Claim Paid Date": pl.Utf8,
    "Dialed Number": pl.Utf8,
    "Days Past Write Off": pl.Int64,
    "Balance": pl.Float64,
    "Contact Type": pl.Utf8,
    "Call Duration": pl.Int64,
    "Talk Time Duration": pl.Int64
}

if "payment_taggings_excel" not in st.session_state:
    st.session_state.payment_taggings_excel = BytesIO()

def concat_df(excel_files, schema):
    # Initialize an empty list to store the DataFrames
    excel_list = []
    
    # Loop through each uploaded file and append the data
    for uploaded_file in excel_files:
        # Read the Excel file into a DataFrame
        df = pl.read_excel(uploaded_file, schema_overrides=schema)
        
        # Append the DataFrame to the list
        excel_list.append(df)
    
    # Concatenate all DataFrames in the list into one DataFrame
    merged_df = pl.concat(excel_list, how="vertical")
    return merged_df

def check_ptp(acc_num, date) -> pl.Struct:
    ptp_dispo = ["PTP - RPC PTP PARTIAL", "PTP - RPC PTP FULL PAYMENT"]
    ptp_status = remark_report.filter((pl.col("Account No.") == acc_num) & (pl.col("PTP Date") == date) & (pl.col("Status").is_in(ptp_dispo)))

    if ptp_status.shape[0] > 0:
        ptp_status = ptp_status.sort("Time", descending=True)
        ptp_status = ptp_status.head(1)

        ptp_result = {
            "Hierachy": "PTP",
            "Date": ptp_status["PTP Date"][0],
            "Status": ptp_status["Status"][0],
            "Remark": ptp_status["Remark"][0],
            "Remark By": ptp_status["Remark By"][0]
        }

        return ptp_result
    else:
        return {
            "Hierachy": None,
            "Date": None,
            "Status": None,
            "Remark": None,
            "Remark By": None
        }
    
def check_rpc(acc_num, date) -> pl.Struct:
    start_date = date - timedelta(days=10)

    rpc_status = remark_report.filter(((pl.col("Date") <= date) & (pl.col("Date") >= start_date)) & (pl.col("Account No.") == acc_num) & ((pl.col("Status").str.starts_with("PTP")) | (pl.col("Status").str.starts_with("POSITIVE CONTACT")) | (pl.col("Status").str.starts_with("PAYMENT"))))


    if rpc_status.shape[0] > 0:
        rpc_status = rpc_status.sort("Time", descending=True)
        rpc_status = rpc_status.head(1)

        rpc_result = {
            "Hierachy": "RPC",
            "Date": rpc_status["Date"][0],
            "Status": rpc_status["Status"][0],
            "Remark": rpc_status["Remark"][0],
            "Remark By": rpc_status["Remark By"][0]
        }

        return rpc_result
    else:
        return {
            "Hierachy": None,
            "Date": None,
            "Status": None,
            "Remark": None,
            "Remark By": None
        }

def check_third_party(acc_num, date) -> pl.Struct:
    start_date = date - timedelta(days=3)

    tpc_status = remark_report.filter(((pl.col("Date") <= date) & (pl.col("Date") >= start_date)) & (pl.col("Account No.") == acc_num) & (pl.col("Status") == "POSITIVE - 3RD PARTY CONTACTED"))


    if tpc_status.shape[0] > 0:
        tpc_status = tpc_status.sort("Time", descending=True)
        tpc_status = tpc_status.head(1)

        tpc_result = {
            "Hierachy": "TPC",
            "Date": tpc_status["Date"][0],
            "Status": tpc_status["Status"][0],
            "Remark": tpc_status["Remark"][0],
            "Remark By": tpc_status["Remark By"][0]
        }

        return tpc_result
    else:
        return {
            "Hierachy": None,
            "Date": None,
            "Status": None,
            "Remark": None,
            "Remark By": None
        }
    
def check_ptp_rpc(acc_num, date) -> dict:
    # First: check PTP
    ptp_result = check_ptp(acc_num, date)
    
    if not all(value is None for value in ptp_result.values()):
        return ptp_result
    
    # Second: check RPC
    rpc_result = check_rpc(acc_num, date)
    
    if not all(value is None for value in rpc_result.values()):
        return rpc_result
    
    # Last: check Third-Party Contact
    tpc_result = check_third_party(acc_num, date)
    return tpc_result
    
with st.form(key="payment_taggings"):
    payments_sheet = st.file_uploader("Payment Sheet", type="xlsx")

    remark_report_files = st.file_uploader("Daily Remark Report", type="xlsx", accept_multiple_files=True)

    if st.form_submit_button(use_container_width=True):
        payments = pl.read_excel(payments_sheet)
        payments = payments.with_columns(
            pl.col("ACC NUMBER").cast(pl.Utf8)
        )

        remark_report = concat_df(remark_report_files, daily_remark_schema)
        remark_report = remark_report.filter(~(pl.col("Remark").str.starts_with("Updates when case reassign")))
        remark_report = remark_report.filter(~(pl.col("Remark").str.contains("New Assignment - OS updated")))
        remark_report = remark_report.filter(~(pl.col("Remark By").is_in(["KCRIO", "SCPEREZ", "RBDELOSREYES", "CTAGLE","NVIRTUDEZ"])))
        

        remark_report = remark_report.with_columns(
            pl.col("PTP Date")
            .str.strptime(pl.Datetime, "%Y-%m-%d %H:%M:%S")
            .dt.date()
            .alias("PTP Date")
        )

        payments = payments.with_columns(
            pl.struct(["ACC NUMBER", "PAYMENT DATE"]).map_elements(lambda x: check_ptp_rpc(x["ACC NUMBER"], x["PAYMENT DATE"]), return_dtype=pl.Struct).alias("RESULT")
        )

        payments = payments.with_columns(
            pl.col("RESULT").struct.unnest()
        )
        payments = payments.drop("RESULT")

        payments.write_excel(
            st.session_state.payment_taggings_excel
        )
        st.session_state.payment_taggings_excel.seek(0)

if st.session_state.payment_taggings_excel.getbuffer().nbytes != 0:
    st.download_button(
        label = "download payment taggings",
        data = st.session_state.payment_taggings_excel.getvalue(),
        file_name = f"maya_payment_taggings.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )