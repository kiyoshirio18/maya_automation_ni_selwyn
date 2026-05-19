import polars as pl
import streamlit as st

st.header("Negosyo SME Dialer Report")

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

def calls_dialed(df: pl.DataFrame):
    dials = df.group_by("Date").agg(
        pl.col("Account No.").len().alias("Dials")
    )
    return dials

def connected_calls(df: pl.DataFrame):
    connected = df.filter(pl.col("Call Duration") > 0)

    connected_count = connected.group_by("Date").agg(
        pl.col("Account No.").len().alias("Connected"),
        pl.col("Debtor").unique().len().alias("Unique Connected")
    )

    return connected_count

def rpc_ptp(df: pl.DataFrame):

    unique = df.group_by("Date").agg(
        pl.col("Debtor").n_unique().alias("UNIQUE COUNT")
    )

    rpc = df.filter((pl.col("Status").str.starts_with("POSITIVE CONTACT")) | (pl.col("Status").str.starts_with("PTP")) | (pl.col("Status").str.starts_with("PAYMENT")))

    rpc_count = rpc.group_by("Date").agg(
        pl.col("Debtor").unique().len().alias("RPC")
    )

    ptp = df.filter((pl.col("Status").str.starts_with("PTP")))

    ptp_count = ptp.group_by("Date").agg(
        pl.col("Debtor").unique().len().alias("PTP")
    )

    rpc_ptp_count = df["Date"].unique().to_frame().join(rpc_count, on="Date", how='left')
    rpc_ptp_count = rpc_ptp_count.join(ptp_count, on="Date", how='left')
    rpc_ptp_count = rpc_ptp_count.join(unique, on="Date", how='left')
    rpc_ptp_count = rpc_ptp_count.with_columns(
        (pl.col("UNIQUE COUNT") - pl.col("RPC")).alias("NON RPC")
    )
    rpc_ptp_count = rpc_ptp_count.select(["Date", "UNIQUE COUNT", "RPC", "NON RPC", "PTP"])

    return rpc_ptp_count

def seconds_to_mmss(seconds):
    total_seconds = int(round(seconds))
    minutes = total_seconds // 60
    remaining_seconds = total_seconds % 60
    return f"{minutes:02}:{remaining_seconds:02}"

def aht(df: pl.DataFrame):
    connected_call = df.filter(pl.col("Call Duration") > 0)

    average_ht = connected_call.group_by("Date").agg(
        pl.col("Call Duration").mean().alias("AHT")
    )

    average_ht = average_ht.with_columns(
        pl.col("AHT").map_elements(seconds_to_mmss, return_dtype = pl.Utf8).alias("AHT")
    )

    return average_ht

def summary(df: pl.DataFrame):
    dials = calls_dialed(df)
    connected = connected_calls(df)
    rpc_ptps = rpc_ptp(df)
    average_ht = aht(df)

    summary = dials.join(connected, on="Date", how="left")
    summary = summary.join(rpc_ptps, on="Date", how="left")
    summary = summary.join(average_ht, on="Date", how="left")

    return summary

with st.form(key="negosyo_sme_dialer"):
    daily_remark_file = st.file_uploader("Daily Remark Report", type="xlsx", accept_multiple_files=True)

    if st.form_submit_button(use_container_width=True):

        daily_remark = concat_df(daily_remark_file, daily_remark_schema)
        daily_remark = daily_remark.with_columns(
            pl.col("Account No.").cast(pl.Utf8)
        )

        negosyo_daily_remark = daily_remark.filter(pl.col("Account No.").str.starts_with("4"))
        sme_daily_remark = daily_remark.filter(pl.col("Account No.").str.starts_with("9"))

        negosyo_summary = summary(negosyo_daily_remark).fill_null(0).sort("Date")
        sme_summary = summary(sme_daily_remark).fill_null(0).sort("Date")

        col1, col2 = st.columns(2)

        col1.markdown("Negosyo :star2:")
        col1.dataframe(negosyo_summary, use_container_width=True)
        col2.markdown("SME :sparkles:")
        col2.dataframe(sme_summary, use_container_width=True)

