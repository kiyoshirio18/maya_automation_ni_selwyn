import polars as pl
import streamlit as st
import json
import xlsxwriter
import pickle
from io import BytesIO

st.header("DCA Monitoring Data")
st.write(f''':green[:green-background[Monthly Remark Report]] :green[:green-background[Daily Remark Report]] :green[:green-background[Placement]]''')

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
        df = pl.read_excel(uploaded_file, engine="openpyxl", schema_overrides=schema)
        
        # Append the DataFrame to the list
        excel_list.append(df)
    
    # Concatenate all DataFrames in the list into one DataFrame
    merged_df = pl.concat(excel_list, how="vertical")
    return merged_df

def get_attempts_summary(volare_remark: pl.DataFrame):
    attempts_count = volare_remark.group_by("Account No.").agg(
        pl.col("Account No.").len().alias("Attempts")
    )

    connected_calls = volare_remark.filter(pl.col("Call Status") == "CONNECTED").group_by("Account No.").agg(
        pl.col("Account No.").len().alias("Connected")
    )

    rpc_calls = volare_remark.filter(
        (pl.col("Call Status") == "CONNECTED") & (pl.col("Status").str.contains("POSITIVE CONTACT"))
    ).group_by("Account No.").agg(
        pl.col("Account No.").len().alias("RPC")
    )

    ptp_calls = volare_remark.filter(
        (pl.col("Call Status") == "CONNECTED") & (pl.col("Status").str.contains("PTP"))
    ).group_by("Account No.").agg(
        pl.col("Account No.").len().alias("PTP")
    )

    join_list = [attempts_count, connected_calls, rpc_calls, ptp_calls]

    attempts_summary = None
    for df in join_list:
        if attempts_summary is None:
            attempts_summary = df
            continue
        attempts_summary = attempts_summary.join(df, on="Account No.", how="left")

    attempts_summary = attempts_summary.fill_null(0)

    return attempts_summary

def get_payments(volare_remark: pl.DataFrame):
    payments = volare_remark.filter(pl.col("Status").str.starts_with("PAYMENT")).sort("Date", descending=True)
    return payments

def format_time(seconds: int):
    """Convert a duration in seconds to a string in the format mm:ss."""
    minutes = seconds // 60
    remaining_seconds = seconds % 60
    return f"{int(minutes):02}:{int(remaining_seconds):02}"

def get_dials_connected(filtered: pl.DataFrame):
    dials = filtered.group_by("Date").agg(
        pl.col("Account No.").len().alias("Dials")
    )

    connected = filtered.filter(pl.col("Call Status") == "CONNECTED").group_by("Date").agg(
        pl.col("Account No.").len().alias("Connected")
    )

    dials_connected = dials.join(connected, on="Date", how="left").sort("Date")
    return dials_connected

def get_call_outcomes(filtered: pl.DataFrame):
    call_outcomes_grouped = filtered.filter(pl.col("Call Status") == "CONNECTED").group_by(["Date", "Call Outcomes"]).agg(
        pl.col("Account No.").len().alias("Count")
    ).sort("Date")

    call_outcomes_pivot = call_outcomes_grouped.pivot(
        on = "Date",
        index = "Call Outcomes",
        values="Count"
    )

    call_outcomes_pivot = call_outcomes_pivot.fill_null(0).sort("Call Outcomes")

    call_outcomes = pl.DataFrame({"Call Outcomes": ["BUSY SIGNAL", "CALL BACK", "CLAIMING FULLY PAID", "DISCONNECTED NUMBER", "DISPUTE", "INVALID NUMBER", "LEFT MESSAGE", "NEW ENDO", "NO ANSWER", "PAYMENT", "PROMISES TO PAY", "REFUSE TO PAY", "UNDER NEGO", "WRONG NUMBER"]})

    call_outcomes = call_outcomes.join(call_outcomes_pivot, on="Call Outcomes", how="left").fill_null(0)

    return call_outcomes

def get_handling_time(filtered: pl.DataFrame):
    handling_time = filtered.filter(
        (pl.col("Call Status") == "CONNECTED") &
        (pl.col("Call Duration").is_not_null()) &
        (pl.col("Call Duration") != 0)
    ).group_by("Date").agg(
        pl.col("Call Duration").mean().alias("Average"),
        pl.col("Call Duration").max().alias("Max"),
        pl.col("Call Duration").min().alias("Min")
    ).sort("Date")

    update_list = ["Average", "Max", "Min"]
    handling_time = handling_time.with_columns(
        pl.col(f"{column}").map_elements(format_time, return_dtype=pl.Utf8).alias(f"{column}") for column in update_list
    )

    return handling_time

with open("./resources/maya_call_outcomes.json", "r") as file:
    call_outcomes_dispo = json.load(file)

if "dialer_summary" not in st.session_state:
    st.session_state.dialer_summary = None

if "attempts_summary" not in st.session_state:
    st.session_state.attempts_summary = None

if "payments_dispo" not in st.session_state:
    st.session_state.payments_dispo = None

if "dca_mon_data" not in st.session_state:
    st.session_state.dca_mon_data = None

if "account_journey" not in st.session_state:
    st.session_state.account_journey = None

placements = ["Maya Credit 121 - 150 DPD", "Maya Credit 181 DPD & UP", "Maya Negosyo Advance 121 - 150 DPD", "Maya Negosyo Advance 181 DPD & UP"]
maya_dispositions = pl.read_csv("./resources/maya_dispositions.csv")

with st.form(key="dca_mon"):
    merged_accounts_file = st.file_uploader("Merged Accounts", type="xlsx")

    monthly_remark_file = st.file_uploader("Monthly Remark Report", type="xlsx", accept_multiple_files=True)

    col1, col2 = st.columns(2, vertical_alignment="bottom")
    start_date = col1.date_input("Start Date")
    end_date = col2.date_input("End Date")

    if st.form_submit_button(use_container_width=True):

        monthly_remark = concat_df(monthly_remark_file, daily_remark_schema)
        daily_remark = monthly_remark.filter((pl.col("Date") >= start_date) & (pl.col("Date") <= end_date))
        merged_accounts = pl.read_excel(merged_accounts_file)

        account_journey = daily_remark.join(maya_dispositions, left_on="Status", right_on="VOLARE STATUS", how="left")
        account_journey = account_journey.with_columns(
            pl.col("HIERARCHY").fill_null(0)
        ).sort("HIERARCHY", descending=True)
        account_journey = account_journey.unique(subset=["Date", "Account No."], keep="first", maintain_order=True)

        daily_remark = daily_remark.join(merged_accounts["ACCOUNT NUMBER", "PLACEMENT"], left_on="Account No.", right_on="ACCOUNT NUMBER", how="left")
        daily_remark = daily_remark.with_columns(
            pl.col("Status").replace(call_outcomes_dispo).alias("Call Outcomes")
        )

        dialer_summary = {}
        for placement in placements:

            if placement not in dialer_summary:
                dialer_summary[placement] = {}

            filtered = daily_remark.filter(
                pl.col("PLACEMENT") == placement
            )
            
            dialer_summary[placement]["dials_connected"] = get_dials_connected(filtered)
            dialer_summary[placement]["call_outcomes"] = get_call_outcomes(filtered)
            dialer_summary[placement]["handling_time"] = get_handling_time(filtered)
        
        st.session_state.dialer_summary = dialer_summary
        st.session_state.attempts_summary = get_attempts_summary(monthly_remark)
        st.session_state.payments_dispo = get_payments(monthly_remark)
        st.session_state.account_journey = account_journey

dca_mon_data = BytesIO()

if all(var is not None for var in [st.session_state.attempts_summary, st.session_state.payments_dispo, st.session_state.account_journey, st.session_state.dialer_summary]):
    with xlsxwriter.Workbook(dca_mon_data) as wb:
        st.session_state.attempts_summary.write_excel(wb, worksheet="ATTEMPTS", column_formats={"Account No.": "0"},autofit=True)
        st.session_state.payments_dispo.write_excel(wb, worksheet="PAYMENTS", column_formats={"Account No.": "0"},autofit=True)
        st.session_state.account_journey.write_excel(wb, worksheet="ACC JOURNEY", dtype_formats={pl.Int64: "0", pl.Date: "mm/dd/yyyy"})

        for placement in placements:
            pos1 = 1
            st.session_state.dialer_summary[placement]["dials_connected"].write_excel(
                wb, 
                worksheet=placement.lstrip("Maya "),
                position=(1, pos1),
                table_style="Table Style Light 8",
                dtype_formats={pl.Int64: "0", pl.Date: "mm/dd/yyyy"},
                autofit=True
            )
            pos2 = pos1 + st.session_state.dialer_summary[placement]["dials_connected"].shape[1] + 1
            st.session_state.dialer_summary[placement]["call_outcomes"].write_excel(wb,
                worksheet=placement.lstrip("Maya "),
                position=(1, pos2),
                table_style="Table Style Light 8",
                dtype_formats={pl.Int64: "0", pl.Date: "mm/dd/yyyy"},
                autofit=True
            )
            pos3 = pos2 + st.session_state.dialer_summary[placement]["call_outcomes"].shape[1] + 1
            st.session_state.dialer_summary[placement]["handling_time"].write_excel(wb,
                worksheet=placement.lstrip("Maya "),
                position=(1, pos3),
                table_style="Table Style Light 8",
                dtype_formats={pl.Int64: "0", pl.Date: "mm/dd/yyyy"},
                autofit=True
            )
        dca_mon_data.seek(0)
    st.download_button(
        label = "Download DCA Monitoring Data",
        data = dca_mon_data.getvalue(),
        file_name = f"maya_dca_monitoring_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )