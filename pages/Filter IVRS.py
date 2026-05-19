import streamlit as st
import polars as pl
from io import BytesIO

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

if "remark_df" not in st.session_state:
    st.session_state.remark_df = None

if "filtered_df" not in st.session_state:
    st.session_state.filtered_df = None

if "date" not in st.session_state:
    st.session_state.dispute_date = None

if "status_list" not in st.session_state:
    st.session_state.status_list = None

if "filter_status" not in st.session_state:
    st.session_state.filter_status = None

def concat_df(excel_files, schema) -> pl.DataFrame:
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

st.header("Filter IVRS")

with st.form(key="ibrs"):
    daily_remark_file = st.file_uploader("Daily Remark Report", type="xlsx", accept_multiple_files=True)
    
    # Only reprocess when the form is submitted
    if st.form_submit_button(use_container_width=True):
        if daily_remark_file:
            # Only process if new file is uploaded
            daily_remark = concat_df(daily_remark_file, daily_remark_schema)
            st.session_state.remark_df = daily_remark

# Make sure to store the unique status values in session state only once
if daily_remark is not None:
    st.session_state.filtered_df = daily_remark.filter(pl.col("Remark").str.contains("Broadcast"))

if st.session_state.filtered_df is not None:
    filtered_xlsx = BytesIO()
    st.session_state.filtered_df.write_excel(
        filtered_xlsx,
        autofit = True
    )
    filtered_xlsx.seek(0)
    st.session_state.date = st.session_state.filtered_df["Date"].max()


    if filtered_xlsx is not None:
        st.download_button(
            label = "Download Filtered IVRS Statuses",
            data = filtered_xlsx.getvalue(),
            file_name = f"maya_ivrs_{st.session_state.date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
