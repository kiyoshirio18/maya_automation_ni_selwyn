import streamlit as st
import polars as pl
import pickle
from io import BytesIO

st.header("Daily Call Logs")


if "daily_remark" not in st.session_state:
    st.session_state.daily_remark = None

if "call_logs" not in st.session_state:
    st.session_state.call_logs = None

if "daily_remark_file" not in st.session_state:
    st.session_state.daily_remark_file = None

if "call_logs_file" not in st.session_state:
    st.session_state.call_logs_file = None

with open("./resources/daily_remark_schema.pkl", "rb") as file:
    daily_remark_schema = pickle.load(file)

daily_remark_schema = {
    "S.No": pl.Int64,
    "Date": pl.Utf8,
    "Time": pl.Utf8,
    "Debtor": pl.Utf8,
    "Account No.": pl.Int64,
    "Card No.": pl.Utf8,
    "Service No.": pl.Utf8,
    "DPD": pl.Int64,
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
    "PTP Amount": pl.Utf8,
    "Next Call": pl.Utf8,
    "PTP Date": pl.Utf8,
    "Claim Paid Amount": pl.Utf8,
    "Claim Paid Date": pl.Utf8,
    "Dialed Number": pl.Utf8,
    "Days Past Write Off": pl.Utf8,
    "Balance": pl.Utf8,
    "Contact Type": pl.Utf8,
    "Cycle": pl.Utf8,
    "Old IC": pl.Utf8,
    "I.C Issue Date": pl.Utf8,
    "Bank Code": pl.Utf8,
    "Over Limit Amount": pl.Float64,
    "Min Payment": pl.Float64,
    "Due Date": pl.Utf8,
    "Monthly Installment": pl.Float64,
    "30 Days": pl.Float64,
    "MIA": pl.Float64,
    "Area": pl.Utf8,
    "Debtor ID": pl.Int64,
    "Call Duration": pl.Utf8,
    "Talk Time Duration": pl.Utf8,
    "Black Case No.": pl.Utf8,
    "Red Case No.": pl.Utf8,
    "Court Name": pl.Utf8,
    "Lawyer": pl.Utf8,
    "Legal Stage": pl.Utf8,
    "Legal Status": pl.Utf8,
    "Next Legal Follow up": pl.Utf8
}

def time_to_seconds(time_str: str):
    hours, minutes, seconds = map(int, time_str.split(":"))
    return hours * 3600 + minutes * 60 + seconds

dispo_list = pl.read_csv("./resources/maya_dispositions.csv")

with st.form(key="daily_call_logs"):

    tx_status_file = st.file_uploader("Volare Texxen Daily Remark", type="xlsx")

    if st.form_submit_button(use_container_width=True):

        tx_status = pl.read_excel(tx_status_file, engine="openpyxl", schema_overrides=daily_remark_schema)
        tx_status = tx_status.drop(["Black Case No.", "Red Case No.", "Court Name", "Lawyer", "Legal Stage", "Legal Status", "Next Legal Follow up", "Old IC", "I.C Issue Date", "Bank Code", "Over Limit Amount", "Min Payment", "Due Date", "Monthly Installment", "30 Days", "MIA", "Area", "Debtor ID"])

        tx_status = tx_status.with_columns(
            pl.col("PTP Amount").str.replace_all(",", "").cast(pl.Float64).alias("PTP Amount"),
            pl.col("Claim Paid Amount").str.replace_all(",", "").cast(pl.Float64).alias("Claim Paid Amount"),
            pl.col("Balance").str.replace_all(",", "").cast(pl.Float64).alias("Balance"),
            pl.col("Date").str.strptime(pl.Date, "%Y-%m-%d %H:%M:%S").alias("Date"),
            pl.col("Time").str.strptime(pl.Time, "%I:%M:%S %p").alias("Time"),
            pl.lit(None).alias("Reason For Default"),
            pl.col("Call Duration").map_elements(time_to_seconds, return_dtype=pl.Int64).alias("Call Duration"),
            pl.col("Talk Time Duration").map_elements(time_to_seconds, return_dtype=pl.Int64).alias("Talk Time Duration")
        )

        tx_status = tx_status.with_columns(
            (pl.col("Date").dt.date().cast(pl.Datetime) + pl.col("Time").cast(pl.Duration)).alias("Time")
        )


        exclude_status = [None, "ABORT", "BP", "NEW", "REACTIVE", "FS", "PP"]

        tx_status = tx_status.filter(~(pl.col("Status").is_in(exclude_status)))

        tx_status = tx_status.with_columns(
            pl.arange(1, tx_status.height + 1).alias("S.No")
        )

        volare_dispo = dict(zip(dispo_list["VOLARE STATUS"].to_list(), dispo_list["PROPOSED DISPOSITION"].to_list()))

        call_logs = tx_status.with_columns(
            pl.col("Status").replace_strict(volare_dispo, default=None).alias("Status")
        )

        call_logs = call_logs.filter(pl.col("Status").is_not_null())

        st.session_state.daily_remark = tx_status.select(["S.No", "Date", "Time", "Debtor", "Account No.", "Cycle", "Card No.", "Service No.", "DPD", "Reason For Default", "Call Status", "Status", "Remark", "Remark By", "Remark Type", "Field Visit Date", "Collector", "Client", "Product Description", "Product Type", "Batch No", "Account Type", "Relation", "PTP Amount", "Next Call", "PTP Date", "Claim Paid Amount", "Claim Paid Date", "Dialed Number", "Days Past Write Off", "Balance", "Contact Type", "Call Duration", "Talk Time Duration"])
        st.session_state.call_logs = call_logs

        st.session_state.daily_remark_file = BytesIO()
        st.session_state.daily_remark.write_excel(
            st.session_state.daily_remark_file,
            dtype_formats={
                pl.Int8: "0",
                pl.Int16: "0",
                pl.Int32: "0",
                pl.Int64: "0",
                pl.Float32: "0.00",
                pl.Float64: "0.00",
                pl.Date: "mm/dd/yyyy",
                pl.Datetime: "mm/dd/yyyy hh:mm:ss",
                pl.Time: "hh:mm:ss"
            },
            autofit=True
        )

        st.session_state.call_logs_file = BytesIO()
        st.session_state.call_logs.write_excel(
            st.session_state.call_logs_file,
            dtype_formats={
                pl.Int8: "0",
                pl.Int16: "0",
                pl.Int32: "0",
                pl.Int64: "0",
                pl.Float32: "0",
                pl.Float64: "0",
                pl.Date: "mm/dd/yyyy",
                pl.Datetime: "mm/dd/yyyy hh:mm:ss"
            },
            autofit=True
        )

if st.session_state.daily_remark_file is not None:
    st.session_state.daily_remark_file.seek(0)

    st.download_button(
        label = f'''Download Daily Remark''',
        data = st.session_state.daily_remark_file,
        file_name = f"maya_daily_remark_{st.session_state.daily_remark["Date"].max()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if st.session_state.call_logs_file is not None:
    st.session_state.call_logs_file.seek(0)

    st.download_button(
        label = f'''Download Daily Call Logs''',
        data = st.session_state.call_logs_file,
        file_name = f"maya_call_logs_{st.session_state.call_logs["Date"].max()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )