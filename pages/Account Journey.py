import polars as pl
import streamlit as st
from io import BytesIO

st.header("Account Journey")
maya_dispositions = pl.read_csv("./resources/maya_dispositions.csv")

daily_remark_schema = {
    "S.No": pl.Int64,
    "Date": pl.Date,
    "Time": pl.Datetime,
    "Debtor": pl.Utf8,
    "Account No.": pl.Int64,
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

if "account_journey" not in st.session_state:
    st.session_state.account_journey = None

with st.form(key="acc_journey"):
    daily_remark_file = st.file_uploader("Daily Remark Report", type="xlsx", accept_multiple_files=True)

    if st.form_submit_button(use_container_width=True):

        daily_remark = concat_df(daily_remark_file, daily_remark_schema)

        account_journey = daily_remark.join(maya_dispositions, left_on="Status", right_on="VOLARE STATUS", how="left")
        account_journey = account_journey.with_columns(
            pl.col("HIERARCHY").fill_null(0)
        ).sort("HIERARCHY", descending=True)
        account_journey = account_journey.unique(subset=["Date", "Account No."], keep="first", maintain_order=True)\

        st.session_state.account_journey = account_journey

if st.session_state.account_journey is not None:
    account_journey_bytes = BytesIO()
    st.session_state.account_journey.write_excel(account_journey_bytes, worksheet="ACC JOURNEY", dtype_formats={pl.Int64: "0", pl.Date: "mm/dd/yyyy"})
    account_journey_bytes.seek(0)

    st.download_button(
        label = "maya account journey",
        data = account_journey_bytes.getvalue(),
        file_name = f"maya_dca_monitoring_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )