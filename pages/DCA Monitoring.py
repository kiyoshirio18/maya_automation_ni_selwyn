import polars as pl
import streamlit as st
import xlsxwriter
from datetime import timedelta
from io import BytesIO
from typing import List


st.header("DCA Monitoring Data")
st.write(
    ":green[:green-background[Monthly Remark Report]] "
    ":green[:green-background[Daily Remark Report]] "
    ":green[:green-background[Placement]]"
)

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
    "Talk Time Duration": pl.Int64,
}

placements = [
    "MC 121-150 DPD",
    "MC 181DPD UP",
    "Maya Negosyo 121DPD",
    "Maya Negosyo 181DPD",
]

placement_agents = {
    "MC 121-150 DPD": [
        "April Angela Dioso",
        "April Mae Boholst",
        "Danielle Mapanao",
        "Jashae Francisco",
        "Joanna Cristine Soverano",
        "Karl Gabriel Villanueva",
        "Kris Mhel Vargas",
        "Nick Andrew Lactaotao",
        "Roxanne Jane Luay",
        "Vhryx Jewelz Macaraig",
    ],
    "MC 181DPD UP": [
        "April Angela Dioso",
        "April Mae Boholst",
        "Danielle Mapanao",
        "Jashae Francisco",
        "Joanna Cristine Soverano",
        "Karl Gabriel Villanueva",
        "Kris Mhel Vargas",
        "Nick Andrew Lactaotao",
        "Roxanne Jane Luay",
        "Vhryx Jewelz Macaraig",
    ],
    "Maya Negosyo 121DPD": [
        "Karl Gabriel Villanueva",
        "Shania Argulla Rosete",
        "Zyrille Bermudez",
    ],
    "Maya Negosyo 181DPD": [
        "Karl Gabriel Villanueva",
        "Shania Argulla Rosete",
        "Zyrille Bermudez",
    ],
}


def safe_sheet_name(name: str) -> str:
    cleaned = name.replace("/", "-").replace("\\", "-").replace("[", "(").replace("]", ")")
    return cleaned[:31]


def col_letter(index: int) -> str:
    letters = ""
    while index >= 0:
        letters = chr(index % 26 + 65) + letters
        index = index // 26 - 1
    return letters


def concat_df(excel_files, schema):
    excel_list = []

    for uploaded_file in excel_files:
        df = pl.read_excel(uploaded_file, engine="openpyxl", schema_overrides=schema)
        excel_list.append(df)

    if not excel_list:
        return pl.DataFrame()

    return pl.concat(excel_list, how="vertical")


def merge_df(df_list: List[pl.DataFrame]) -> pl.DataFrame:
    if not df_list:
        raise ValueError("df_list is empty.")

    merged = df_list[0]
    join_key = merged.columns[0]

    for df in df_list[1:]:
        merged = merged.join(df, on=join_key, how="left")

    return merged


def format_seconds(seconds: int | float) -> str:
    if seconds is None:
        return "00:00"
    minutes, secs = divmod(int(seconds), 60)
    return f"{minutes:02d}:{secs:02d}"


def dialer_report(placement, df: pl.DataFrame):
    connected = df.filter(pl.col("Call Duration") > 0)

    result = connected.group_by("AGENT NAME").agg(
        [
            pl.col("Account No.").len().alias("Total Connected Calls"),
            (
                (
                    pl.col("CALL OUTCOMES").str.starts_with("RPC")
                    | pl.col("CALL OUTCOMES").str.starts_with("RIGHT PARTY")
                )
                & ~pl.col("CALL OUTCOMES").str.contains("Right Party - Promise to Pay")
            ).sum().alias("Right Party Contact"),
            (pl.col("CALL OUTCOMES") == "RIGHT PARTY - PROMISE TO PAY").sum().alias("Promise to Pay"),
            pl.col("CALL OUTCOMES").str.starts_with("THIRD PARTY").sum().alias("Third Party Contact"),
            (pl.col("CALL OUTCOMES") == "THIRD PARTY - REFUSED TO RELAY MESSAGE").sum().alias("Relative (Refused)"),
            (pl.col("CALL OUTCOMES") == "THIRD PARTY - DOES NOT KNOW BORROWER").sum().alias("Customer Unknown"),
            (pl.col("CALL OUTCOMES") == "NO ANSWER FROM THE USER").sum().alias("No Answer"),
            (pl.col("CALL OUTCOMES") == "DROPPED CALL").sum().alias("Dropped Call"),
            (pl.col("CALL OUTCOMES") == "NOT IN SERVICE").sum().alias("Uncontacted (NIS, OOCA)"),
            (pl.col("CALL OUTCOMES") == "THIRD PARTY - WRONG NUMBER").sum().alias("Incorrect Number"),
            (pl.col("CALL OUTCOMES") == "CALL REJECTED").sum().alias("Call Rejected"),
            (pl.col("CALL OUTCOMES") == "CALL ENDED").sum().alias("Call Ended"),
            (pl.col("CALL OUTCOMES") == "BUSY").sum().alias("Busy"),
        ]
    )

    agent_list = pl.DataFrame({"AGENT NAME": placement_agents.get(placement, [])})
    if agent_list.height == 0:
        return result

    result = agent_list.join(result, on="AGENT NAME", how="left")
    return result.fill_null(0)


def vb_count(day, df: pl.DataFrame):
    day = day.strftime("%Y-%m-%d")
    broadcast = df.filter(pl.col("Remark").str.contains("broadcast"))

    count = broadcast.height
    connected = broadcast.filter((pl.col("Status") == "PU") | (pl.col("Status") == "PM")).height
    not_connected = count - connected

    return pl.DataFrame({
        "vb": ["count", "connected", "not_connected"],
        day: [count, connected, not_connected],
    })


def call_outcomes(day, df: pl.DataFrame):
    day = day.strftime("%Y-%m-%d")
    connected = df.filter(pl.col("Call Duration") > 0)

    dispo_list = [
        "RIGHT PARTY - PROMISE TO PAY",
        "RIGHT PARTY - NO PROMISE TO PAY",
        "RIGHT PARTY - PAID TODAY",
        "RIGHT PARTY - ALREADY PAID PRIOR",
        "RIGHT PARTY - REQUESTED RESTRUCTURING",
        "RIGHT PARTY - DISPUTING THE ACCOUNT",
        "RIGHT PARTY - FINANCIAL HARDSHIP",
        "RIGHT PARTY - CALL BACK REQUESTED",
        "RIGHT PARTY - REFUSED TO TALK",
        "RIGHT PARTY - NO INTENTION OF PAYING",
        "RIGHT PARTY - CALL REJECTED",
        "RIGHT PARTY - OUT OF THE COUNTRY/TOWN",
        "RPC - INSOLVENCY/BANKRUPTCY",
        "RPC - ENDED MIDDLE OF THE CALL",
        "RPC - CANNOT PAY DUE TO MAYA APP ERROR",
        "RIGHT PARTY - UNAUTHORIZED TRANSACTION",
        "THIRD PARTY - CONFIRMED BORROWER AVAILABILITY",
        "THIRD PARTY - NOT AVAILABLE",
        "THIRD PARTY - MESSAGE RELAYED",
        "THIRD PARTY - BORROWER WILL PAY SOON",
        "THIRD PARTY - DOES NOT KNOW BORROWER",
        "THIRD PARTY - WRONG NUMBER",
        "THIRD PARTY - REFUSED TO RELAY MESSAGE",
        "THIRD PARTY - CALL REJECTED",
        "THIRD PARTY - OUT OF THE COUNTRY/TOWN",
        "THIRD PARTY - CLIENT DECEASED",
        "NO ANSWER FROM THE USER",
        "CALL ENDED",
        "DROPPED CALL",
        "BUSY",
        "NOT IN SERVICE",
        "OUT OF COVERAGE",
        "CALL REJECTED",
        "INCORRECT NUMBER",
        "Pull Out",
    ]

    result = pl.DataFrame({"Call Outcomes": dispo_list})
    count_per_dispo = connected.group_by("CALL OUTCOMES").agg(pl.col("Account No.").len().alias(day))

    result = result.join(count_per_dispo, left_on="Call Outcomes", right_on="CALL OUTCOMES", how="left")
    return result.with_columns(pl.col(day).fill_null(0))


def call_duration(day, df: pl.DataFrame) -> pl.DataFrame:
    day = day.strftime("%Y-%m-%d")
    connected = df.filter(pl.col("Call Duration") > 0)

    if connected.height == 0:
        return pl.DataFrame({"call duration": ["average", "max", "min"], day: ["00:00", "00:00", "00:00"]})

    average = connected["Call Duration"].mean()
    max_val = connected["Call Duration"].max()
    min_val = connected["Call Duration"].min()

    return pl.DataFrame({
        "call duration": ["average", "max", "min"],
        day: [format_seconds(average), format_seconds(max_val), format_seconds(min_val)],
    })


def dialer_report_summary(day, df: pl.DataFrame):
    if df.is_empty():
        return pl.DataFrame({
            "date": [day],
            "connected": [0],
            "rpc": [0],
            "tpc": [0],
            "no_answer": [0],
            "uncon": [0],
        })

    connected_df = df.filter(pl.col("Call Duration") > 0)
    connected = connected_df.height or 0

    rpc = connected_df.filter(
        (pl.col("CALL OUTCOMES").str.contains("RPC -"))
        | (pl.col("CALL OUTCOMES").str.contains("PTP"))
    ).height or 0

    tpc = connected_df.filter(pl.col("CALL OUTCOMES").str.contains("TP -")).height or 0
    no_answer = connected_df.filter(pl.col("CALL OUTCOMES") == "NO ANSWER FROM THE USER").height or 0
    uncon = max(connected - (rpc + tpc + no_answer), 0)

    return pl.DataFrame({
        "date": [day],
        "connected": [connected],
        "rpc": [rpc],
        "tpc": [tpc],
        "no_answer": [no_answer],
        "uncon": [uncon],
    })


def build_report_workbook(remark_report: pl.DataFrame, start_date, end_date) -> BytesIO:
    workbook_bytes = BytesIO()
    workbook = xlsxwriter.Workbook(workbook_bytes)

    # Keep track of created worksheets to avoid recreating them
    created_worksheets: dict = {}

    def get_or_add_worksheet(name: str):
        if name in created_worksheets:
            return created_worksheets[name]
        ws = workbook.add_worksheet(name)
        created_worksheets[name] = ws
        return ws

    def write_df_manual(sheet_name: str, start_col: int, df: pl.DataFrame):
        ws = get_or_add_worksheet(sheet_name)
        # write header
        for j, col_name in enumerate(df.columns):
            ws.write(0, start_col + j, col_name)
        # write rows
        for i, row in enumerate(df.rows()):
            for j, val in enumerate(row):
                ws.write(1 + i, start_col + j, val)

    for placement in placements:
        per_placement = remark_report.filter(pl.col("Cycle") == placement)

        call_outcomes_list = []
        call_duration_list = []
        vb_count_list = []
        dialer_summary_list = []

        current_date = start_date
        while current_date <= end_date:
            daily_remark = per_placement.filter(pl.col("Date") == current_date)

            call_outcomes_result = call_outcomes(current_date, daily_remark)
            dialer_report_result = dialer_report(placement, daily_remark)
            ivrs_result = vb_count(current_date, daily_remark)
            call_duration_result = call_duration(current_date, daily_remark)
            summary_result = dialer_report_summary(current_date, daily_remark)

            call_outcomes_list.append(call_outcomes_result)
            call_duration_list.append(call_duration_result)
            vb_count_list.append(ivrs_result)
            dialer_summary_list.append(summary_result)

            sheet_name = safe_sheet_name(f"{placement[:12]}_{current_date.strftime('%m%d%y')}")
            col_offset = 0

            # write per-day tables manually to avoid xlsxwriter table overlap
            write_df_manual(sheet_name, col_offset, call_outcomes_result)
            col_offset += len(call_outcomes_result.columns) + 1
            write_df_manual(sheet_name, col_offset, dialer_report_result)
            col_offset += len(dialer_report_result.columns) + 1
            write_df_manual(sheet_name, col_offset, ivrs_result)
            col_offset += len(ivrs_result.columns) + 1
            write_df_manual(sheet_name, col_offset, call_duration_result)

            current_date += timedelta(days=1)

        merge_df(call_outcomes_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} CALL OUTCOMES"),
            autofit=True,
        )
        merge_df(call_duration_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} CALL DURATION"),
            autofit=True,
        )
        merge_df(vb_count_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} IVRS"),
            autofit=True,
        )
        pl.concat(dialer_summary_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} SUMMARY"),
            autofit=True,
        )

    workbook.close()
    workbook_bytes.seek(0)
    return workbook_bytes


def build_report_workbooks(remark_report: pl.DataFrame, start_date, end_date) -> dict:
    """Build a separate workbook BytesIO for each placement and return a mapping
    of normalized filename -> BytesIO.
    """
    name_map = {
        "Maya Negosyo 121DPD": "Maya_Negosyo_121DPD",
        "Maya Negosyo 181DPD": "Maya_Negosyo_181DPD",
        "MC 121-150 DPD": "MC_121-150_DPD",
        "MC 181DPD UP": "MC_181DPD_UP",
    }

    results = {}
    for placement in placements:
        per_placement = remark_report.filter(pl.col("Cycle") == placement)

        workbook_bytes = BytesIO()
        workbook = xlsxwriter.Workbook(workbook_bytes)

        created_worksheets: dict = {}

        def get_or_add_worksheet(name: str):
            if name in created_worksheets:
                return created_worksheets[name]
            ws = workbook.add_worksheet(name)
            created_worksheets[name] = ws
            return ws

        def write_df_manual(sheet_name: str, start_col: int, df: pl.DataFrame):
            ws = get_or_add_worksheet(sheet_name)
            for j, col_name in enumerate(df.columns):
                ws.write(0, start_col + j, col_name)
            for i, row in enumerate(df.rows()):
                for j, val in enumerate(row):
                    ws.write(1 + i, start_col + j, val)

        call_outcomes_list = []
        call_duration_list = []
        vb_count_list = []
        dialer_summary_list = []

        current_date = start_date
        while current_date <= end_date:
            daily_remark = per_placement.filter(pl.col("Date") == current_date)

            call_outcomes_result = call_outcomes(current_date, daily_remark)
            dialer_report_result = dialer_report(placement, daily_remark)
            ivrs_result = vb_count(current_date, daily_remark)
            call_duration_result = call_duration(current_date, daily_remark)
            summary_result = dialer_report_summary(current_date, daily_remark)

            call_outcomes_list.append(call_outcomes_result)
            call_duration_list.append(call_duration_result)
            vb_count_list.append(ivrs_result)
            dialer_summary_list.append(summary_result)

            sheet_name = safe_sheet_name(f"{placement[:12]}_{current_date.strftime('%m%d%y')}")
            col_offset = 0

            write_df_manual(sheet_name, col_offset, call_outcomes_result)
            col_offset += len(call_outcomes_result.columns) + 1
            write_df_manual(sheet_name, col_offset, dialer_report_result)
            col_offset += len(dialer_report_result.columns) + 1
            write_df_manual(sheet_name, col_offset, ivrs_result)
            col_offset += len(ivrs_result.columns) + 1
            write_df_manual(sheet_name, col_offset, call_duration_result)

            current_date += timedelta(days=1)

        # write merged sheets
        merge_df(call_outcomes_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} CALL OUTCOMES"),
            autofit=True,
        )
        merge_df(call_duration_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} CALL DURATION"),
            autofit=True,
        )
        merge_df(vb_count_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} IVRS"),
            autofit=True,
        )
        pl.concat(dialer_summary_list).write_excel(
            workbook,
            worksheet=safe_sheet_name(f"{placement[:20]} SUMMARY"),
            autofit=True,
        )

        workbook.close()
        workbook_bytes.seek(0)

        fname = name_map.get(placement, placement.replace(" ", "_").replace("/", "-"))
        results[fname] = workbook_bytes

    return results


dispositions = pl.read_csv("./resources/maya_dispositions.csv")
agent_ref = pl.read_csv("./resources/agent_ref.csv")

if "dca_monitoring_files" not in st.session_state:
    st.session_state.dca_monitoring_files = None

with st.form(key="dca_mon"):
    daily_remark_file = st.file_uploader("Daily Remark Report", type="xlsx", accept_multiple_files=True)

    col1, col2 = st.columns(2, vertical_alignment="bottom")
    start_date = col1.date_input("Start Date")
    end_date = col2.date_input("End Date")

    submitted = st.form_submit_button(use_container_width=True)

if submitted:
    if not daily_remark_file:
        st.error("Upload at least one Daily Remark Report file.")
    elif end_date < start_date:
        st.error("End Date must be on or after Start Date.")
    else:
        remark_report = concat_df(daily_remark_file, daily_remark_schema)
        remark_report = remark_report.filter(~pl.col("Remark").str.starts_with("Updates when case reassign"))
        remark_report = remark_report.filter(~pl.col("Remark").str.contains("New Assignment - OS updated"))
        remark_report = remark_report.join(dispositions, left_on="Status", right_on="VOLARE STATUS", how="left")
        remark_report = remark_report.join(agent_ref, left_on="Remark By", right_on="VOLARE USERNAME", how="left")

        # Normalize disposition column names so downstream code can rely on
        # a single `CALL OUTCOMES` column regardless of source file naming.
        dispo_candidates = [
            "CALL OUTCOMES",
            "Call Status",
            "PROPOSED DISPOSITION",
            "PROPOSED_DISPOSITION",
            "PROPOSED DISPOSITION",
        ]
        found = None
        for c in dispo_candidates:
            if c in remark_report.columns:
                found = c
                break

        if found:
            if found != "CALL OUTCOMES":
                remark_report = remark_report.with_columns(pl.col(found).alias("CALL OUTCOMES"))
        else:
            # If no disposition column exists, create an empty one to avoid errors
            remark_report = remark_report.with_columns(pl.lit("").alias("CALL OUTCOMES"))

        st.session_state.dca_monitoring_files = build_report_workbooks(remark_report, start_date, end_date)
        st.success("DCA Monitoring workbooks are ready for download.")
if st.session_state.dca_monitoring_files:
    # present four download buttons in desired order
    order = [
        "Maya_Negosyo_121DPD",
        "Maya_Negosyo_181DPD",
        "MC_121-150_DPD",
        "MC_181DPD_UP",
    ]
    cols = st.columns(4)
    for c, name in zip(cols, order):
        if name in st.session_state.dca_monitoring_files:
            bio = st.session_state.dca_monitoring_files[name]
            bio.seek(0)
            file_label = name + ".xlsx"
            c.download_button(
                label=f"Download {name}",
                data=bio,
                file_name=file_label,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )