import streamlit as st
import polars as pl
import msoffcrypto
from io import BytesIO
from datetime import datetime

st.header("Blasting Import File")
st.write(f''':green[:green-background[Masterfile]]''')

def save_excel(_df: pl.DataFrame) -> bytes:
    output = BytesIO()
    if _df is not None:
        _df.write_excel(
            output, 
            autofit=True,
            column_formats={"ACCOUNT_NUMBER": "0",
                            "RECEIVED DATE": "mm/dd/yyyy",
                            "OB": "0.00"})
    output.seek(0)
    return output.getvalue()

if 'df' not in st.session_state:
    st.session_state.df = None

if "st.session_state.placements" not in st.session_state:
    st.session_state.placements = None

if "others" not in st.session_state:
    st.session_state.others = None

if "filter_options" not in st.session_state:
    st.session_state.filter_options = None

if "final_df" not in st.session_state:
    st.session_state.final_df = None

col1, col2 = st.columns(2, vertical_alignment="bottom")
col1_1, col1_2 = col1.columns(2)

with col1.form(key="blasting"):
    file = st.file_uploader("Upload Password Protected XLSX", type=["xlsx"])
    passkey = st.text_input("Password", type="password")

    if st.form_submit_button("SUBMIT", use_container_width=True):
        if file is not None:
            try:
                excel_decrypted = BytesIO()
                with BytesIO(file.read()) as f:
                    excel_file = msoffcrypto.OfficeFile(f)
                    excel_file.load_key(password = passkey)
                    excel_file.decrypt(excel_decrypted)
                excel_decrypted.seek(0)
                st.session_state.df = pl.read_excel(excel_decrypted, sheet_name='ACTIVE', schema_overrides={
                    'ID_NUMBER': pl.Utf8,
                    'ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE': pl.Utf8,
                    'RECEIVED DATE': pl.Date,
                    'BIRTH_DATE': pl.Date,
                    'ACCOUNT NUMBER': pl.Int64
                })
            except msoffcrypto.exceptions.DecryptionError as e:
                if str(e) == "Unencrypted document" or str(e) == "No key specified":
                    print("Caught DecryptionError with message: Unencrypted document")
                    st.session_state.df = pl.read_excel(file, schema_overrides={
                        'ID_NUMBER': pl.Utf8,
                        'ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE': pl.Utf8,
                        'RECEIVED DATE': pl.Date,
                        'BIRTH_DATE': pl.Date,
                        'ACCOUNT NUMBER': pl.Int64
                    })
                else:
                    print(f"Caught DecryptionError with a different message: {e}")

        st.session_state.placements = pl.Series(st.session_state.df.select(pl.col("PLACEMENT"))).unique().to_list()
        st.session_state.others = pl.Series(st.session_state.df.select(pl.col("FRESH/SPILLOVER"))).unique().to_list()
        st.session_state.options = [f"{placement} | {other}" for placement in st.session_state.placements for other in st.session_state.others]
        st.session_state.filter_options = (
            st.session_state.df.select((pl.col("PLACEMENT").cast(str) + " | " + pl.col("FRESH/SPILLOVER")).alias("concatenated"))
            .unique()
            .to_series()
            .to_list()
        )

if st.session_state.df is not None:
    summary = st.session_state.df.group_by(["PLACEMENT", "FRESH/SPILLOVER"]).len()
    summary = summary.sort("PLACEMENT", "FRESH/SPILLOVER")
    summary = summary.rename({"len":"Total Accounts"})

    col2.dataframe(summary, use_container_width=True, height=474)

    with st.container():
        sorted_options = sorted(st.session_state.filter_options)
        filters = col1.multiselect("FILTER", sorted_options)

        col_1, col_2 = col1.columns(2)
        if col1.button("FILTER", use_container_width=True):
            if filters:
                filter_df = pl.DataFrame([])

                with pl.Config(
                    tbl_cell_numeric_alignment="RIGHT",
                    thousands_separator=False
                ):

                    for filter_option in filters:
                        placement, fs = filter_option.split(" | ")
                        filtered_rows = st.session_state.df.filter((pl.col("PLACEMENT") == placement) & (pl.col("FRESH/SPILLOVER") == fs))
                        filter_df = filter_df.vstack(filtered_rows)

                    st.session_state.final_df = filter_df.select(["CHCODE", "ACCOUNT_NUMBER", "NAME", "OB", "RECEIVED DATE", "MOBILE PROPER", "EMAIL_ADDRESS", "PLACEMENT", "DPD_", "DPD BUCKET"])
            else:
                st.session_state.final_df = st.session_state.df.select(["CHCODE", "ACCOUNT_NUMBER", "NAME", "OB", "RECEIVED DATE", "MOBILE PROPER", "EMAIL_ADDRESS", "PLACEMENT", "DPD_", "DPD BUCKET"])

if st.session_state.final_df is not None:
    col1, col2 = st.columns(2)
    col1.dataframe(st.session_state.final_df, use_container_width=True)
    col2.dataframe(pl.Series(st.session_state.final_df.select("PLACEMENT")).value_counts(), use_container_width=True)

    now = datetime.now()
    st.download_button(
        label = "Download .xlsx",
        data = save_excel(st.session_state.final_df),
        file_name = f"maya_blasting_{now.strftime("%m%d%y")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )