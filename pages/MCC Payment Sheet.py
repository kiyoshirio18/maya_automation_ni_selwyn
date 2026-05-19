import polars as pl
import streamlit as st
from io import BytesIO

st.header("Payments Sheet")
st.write(f''':green[:green-background[Payment Files]]''')

if "payments_sheet" not in st.session_state:
    st.session_state.payments_sheet = pl.DataFrame()

def save_excel(_df: pl.DataFrame, formatting) -> bytes:
    output = BytesIO()
    if _df is not None:
        _df.write_excel(
            output, 
            autofit=True,
            column_formats=formatting)
    output.seek(0)
    return output.getvalue()

with st.form(key="payment_sheet"):

    mcc_raw_file = st.file_uploader("Maya CC", type=["xlsx"])

    if st.form_submit_button("SUBMIT", use_container_width=True):

        if mcc_raw_file is not None:
            mcc_payments = pl.read_excel(mcc_raw_file, engine="openpyxl")

            if mcc_payments.schema["NAME"] != pl.Null:
                mcc_payments = mcc_payments.select(["ACCOUNT_NUMBER", "NAME", "CREATION_DATE", "TOTAL_PAYMENT", "TRANSACTION_ID", "PRODUCT_NAME"])
                mcc_payments = mcc_payments.rename({
                "NAME": "FULL_NAME",
                "CREATION_DATE": "CREATED_DATE"
            })
            else:
                mcc_payments = mcc_payments.select(["ACCOUNT_NUMBER", "LAST_NAME", "CREATION_DATE", "TOTAL_PAYMENT", "TRANSACTION_ID", "PRODUCT_NAME"])
                mcc_payments = mcc_payments.rename({
                "LAST_NAME": "FULL_NAME",
                "CREATION_DATE": "CREATED_DATE"
            })

            mayacredit_payments = mcc_payments.with_columns(
                pl.col("ACCOUNT_NUMBER").cast(pl.Int64),
                pl.col("FULL_NAME").cast(pl.Utf8),
                pl.col("CREATED_DATE").cast(pl.Date),
                pl.col("TOTAL_PAYMENT").cast(pl.Float64),
                pl.col("TRANSACTION_ID").cast(pl.Utf8),
                pl.col("PRODUCT_NAME").cast(pl.Utf8)
            )
        
        else:
            mcc_payments = None

        payments_sheet = mcc_payments

        st.session_state.payments_sheet = payments_sheet

if st.session_state.payments_sheet.shape[1] > 0:
    st.dataframe(st.session_state.payments_sheet, use_container_width=True)

    formatting = {
        "ACCOUNT_NUMBER": "0",
        "CREATED_DATE": "mm/dd/yy",
        "TOTAL_PAYMENT": "0.00"
    }

    date = st.session_state.payments_sheet["CREATED_DATE"].max().strftime("%m%d%y")

    st.download_button(
        label = "Download Payments Sheet",
        data = save_excel(st.session_state.payments_sheet, formatting),
        file_name = f"maya_payments_sheet_{date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )