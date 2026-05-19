import streamlit as st
import polars as pl
import pickle
import re
from datetime import datetime, timedelta, time
from io import BytesIO

st.header("LOXON Upload")

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

status_ref = pl.read_excel("./resources/maya_reference.xlsx", sheet_name="STATUS")
agent_ref = pl.read_excel("./resources/maya_reference.xlsx", sheet_name="AGENT")

if "loxon_upload_file" not in st.session_state:
    st.session_state.loxon_upload_file = None
if "file_date" not in st.session_state:
    st.session_state.file_date = None

def add_seconds(start_time: datetime, seconds: int) -> datetime:
    return start_time + timedelta(seconds=seconds)

def seconds_to_time(seconds: int) -> time:
    """Convert seconds to a time object (HH:MM:SS)."""
    hours = (seconds // 3600) % 24
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return time(hour=hours, minute=minutes, second=seconds)

def extract_rfd(text: str) -> str:
    match = re.search(r'RFD: (.*?) \|', text)
    return match.group(1) if match else None

def product_name(input_value):
    input_string = str(input_value)  # Convert the input to a string
    if input_string.startswith('6'):
        return "MayaCredit"
    elif input_string.startswith('4'):
        return "negosyoAdvance"
    elif input_string.startswith('9'):
        return "MAYA_FLEXI_ENTERPRISE_LOAN"
    else:
        return ""
    
def format_phone_number(phone: str) -> str:
    phone = str(phone)
    if phone.startswith('09'):
        return '63' + phone[1:]  # Remove leading 0 and add 63
    elif phone.startswith('9'):
        return '63' + phone
    elif phone.startswith('+63'):
        return phone[1:]  # Remove the '+'
    else:
        return phone
    
def clean_status(df: pl.DataFrame) -> pl.DataFrame:
    """
    Cleans a pandas DataFrame by removing rows based on rules for 
    'remark' and 'status_code' columns.
    """

    # 1. Remove rows where 'remark' contains any of these phrases
    remark_filters = [
        "Updates when case reassign",
        "System Auto Update",
        "New Assignment - OS updated"
    ]

    # Build a regex OR pattern
    remark_pattern = "|".join(remark_filters)

    # Keep rows where remark does NOT contain the unwanted text
    df = df.filter(~(pl.col("Remark").str.contains(remark_pattern)))

    # 2. Remove rows where 'status_code' is any of these values
    status_remove = ["ABORT", "LOCKED", "NEW", "REACTIVE", "BP", "LOCKED"]

    df = df.filter(~(pl.col("Status").is_in(status_remove)))

    # Reset index for cleanliness
    return df

with st.form(key="loxon_ulpoad"):
    merged_accounts_file = st.file_uploader("Merged Accounts", type="xlsx")
    daily_remark_file = st.file_uploader("Daily Remark", type="xlsx")

    if st.form_submit_button(use_container_width=True):
        merged_accounts = pl.read_excel(merged_accounts_file)

        daily_remark = clean_status(pl.read_excel(daily_remark_file, schema_overrides=daily_remark_schema))

        daily_remark = daily_remark.with_columns(
            pl.col("Call Duration").fill_null(0).alias("Call Duration")
        )

        daily_remark = daily_remark.join(status_ref, left_on="Status", right_on="VOLARE STATUS", how="left")
        daily_remark = daily_remark.join(agent_ref["VOLARE USERNAME", "AGENT NAME"], left_on="Remark By", right_on="VOLARE USERNAME", how="left")

        daily_remark = daily_remark.with_columns(
            pl.struct(["Time", "Call Duration"]).map_elements(lambda x: add_seconds(x["Time"], x["Call Duration"]), return_dtype=pl.Datetime).alias("END"),
            pl.col("Remark").map_elements(extract_rfd, return_dtype=pl.Utf8).alias("RFD"),
            pl.col("Call Duration").map_elements(seconds_to_time, return_dtype=pl.Time).alias("duration"),
            pl.col("Account No.").map_elements(product_name, return_dtype=pl.Utf8).alias("product_name")
        ).sort("Hierarchy", descending=True)

        daily_remark = daily_remark.with_columns(
            pl.col("Hierarchy").fill_null(0)
        ).sort("Hierarchy", descending=True)

        daily_remark_unique = daily_remark.unique(subset="Account No.", keep="first", maintain_order=True)

        loxon_upload = daily_remark_unique.select(["Time", "result", "Remark", "Account No.", "Dialed Number", "PTP Date", "PTP Amount", "RFD", "END", "duration", "AGENT NAME", "product_name", "channel"])
        loxon_upload = loxon_upload.with_columns(
            pl.lit(None).alias("outsource_case_id"),
            pl.lit("Madrid").alias("outsource_partner_alias"),
            pl.lit(None).alias("skip_phone_number"),
            pl.lit(None).alias("skip_email"),
            pl.lit(None).alias("nonvoice_template"),
            pl.col("Time").alias("event_datetime_pht"),
            pl.col("Time").alias("datalate_processed_ts_pht"),
            pl.col("Dialed Number").map_elements(format_phone_number, return_dtype=pl.Utf8).alias("Dialed Number")
        )

        loxon_upload = loxon_upload.join(merged_accounts["ACCOUNT NUMBER", "AGENT NAME", "CPM_ID", "ACCOUNT_ID", "MOBILE PROPER"], left_on="Account No.", right_on="ACCOUNT NUMBER", how="left")

        loxon_upload = loxon_upload.rename({
            "Time": "call_start",
            "Remark": "comment",
            "Account No.": "account_number",
            "Dialed Number": "number_contacted",
            "PTP Date": "ptp_date",
            "PTP Amount": "ptp_amount",
            "RFD": "reason_for_delay",
            "END": "call_end",
            "AGENT NAME": "collector_full_name",
            "ACCOUNT_ID": "account_id",
            "CPM_ID": "cpm_id",
            "channel": "communication_channel"
        })

        loxon_upload = loxon_upload.with_columns(
            pl.when((pl.col("number_contacted").is_null()) & (pl.col("communication_channel") == "VOICE")).then(
                pl.col("MOBILE PROPER")
            ).otherwise(
                pl.col("number_contacted")
            ).alias("number_contacted"),
            pl.when(pl.col("collector_full_name").is_null()).then(
                pl.col("AGENT NAME_right")
            ).otherwise(
                pl.col("collector_full_name")
            ).alias("collector_full_name"),
            pl.col("ptp_amount").replace(0.0, None)
        )

        loxon_upload = loxon_upload.with_columns(
            pl.col("number_contacted").map_elements(format_phone_number, return_dtype=pl.Utf8).alias("number_contacted")
        )

        loxon_upload = loxon_upload.select([
            "outsource_case_id", "outsource_partner_alias", "event_datetime_pht", "result", "comment",
            "cpm_id", "account_id", "account_number", "product_name", "communication_channel", "number_contacted",
            "ptp_date", "ptp_amount", "reason_for_delay", "call_start", "call_end", "duration",
            "collector_full_name", "skip_phone_number", "skip_email", "nonvoice_template","datalate_processed_ts_pht"
        ])

        loxon_upload_file = BytesIO()

        loxon_upload.write_excel(
            loxon_upload_file,
            column_formats={
                "account_number": "0",
                "ptp_amount": "0.00",
                "call_start": "hh:mm:ss",
                "call_end": "hh:mm:ss",
                "duration": "hh:mm:ss"
            },
            autofit=True
        )

        loxon_upload_file.seek(0)

        st.session_state.loxon_upload_file = loxon_upload_file
        st.session_state.file_date = loxon_upload["event_datetime_pht"].max().date()

if st.session_state.loxon_upload_file is not None:
    st.download_button(
        "Download Daily LOXON Upload",
        data=st.session_state.loxon_upload_file,
        file_name=f"MADRID_Feedback_{st.session_state.file_date}.xlsx",
        use_container_width=True
    )