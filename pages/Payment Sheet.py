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

    col1, col2, col3 = st.columns(3)

    mayacredit_raw_file = col1.file_uploader("Maya Credit", type=["xlsx"])
    negosyo_raw_file = col2.file_uploader("Negosyo", type=["xlsx"])
    sme_raw_file = col3.file_uploader("SME", type=["xlsx"])

    if st.form_submit_button("SUBMIT", use_container_width=True):

        if mayacredit_raw_file is not None:
            mayacredit_payments = pl.read_excel(mayacredit_raw_file)

            if mayacredit_payments.schema["ACCOUNT_NAME"] != pl.Null:
                mayacredit_payments = mayacredit_payments.select(["ACCOUNT_NUMBER", "ACCOUNT_NAME", "COLLECTION_DATE", "TOTAL_REPAYMENT", "TRANSACTION_ID", "PRODUCT_TYPE"])
                mayacredit_payments = mayacredit_payments.rename({
                "ACCOUNT_NAME": "FULL_NAME",
                "COLLECTION_DATE": "CREATED_DATE",
                "TOTAL_REPAYMENT": "TOTAL_PAYMENT",
                "PRODUCT_TYPE": "PRODUCT_NAME"
            })
            elif mayacredit_payments.schema["ACCOUNT_NAME"] != pl.Null:
                mayacredit_payments = mayacredit_payments.select(["ACCOUNT_NUMBER", "ACCOUNT_NAME", "COLLECTION_DATE", "TOTAL_REPAYMENT", "TRANSACTION_ID", "PRODUCT_TYPE"])
                mayacredit_payments = mayacredit_payments.rename({
                "ACCOUNT_NAME": "FULL_NAME",
                "COLLECTION_DATE": "CREATED_DATE",
                "TOTAL_REPAYMENT": "TOTAL_PAYMENT",
                "PRODUCT_TYPE": "PRODUCT_NAME"
            })

            mayacredit_payments = mayacredit_payments.with_columns(
                pl.col("ACCOUNT_NUMBER").cast(pl.Int64),
                pl.col("FULL_NAME").cast(pl.Utf8),
                pl.col("CREATED_DATE").cast(pl.Date),
                pl.col("TOTAL_PAYMENT").cast(pl.Float64),
                pl.col("TRANSACTION_ID").cast(pl.Utf8),
                pl.col("PRODUCT_NAME").cast(pl.Utf8)
            )
        
        else:
            mayacredit_payments = None

        if negosyo_raw_file is not None:
            negosyo_payments = pl.read_excel(negosyo_raw_file)

            negosyo_payments = negosyo_payments.with_columns(
                (pl.col("FIRST_NAME") + " " + pl.col("LAST_NAME")).alias("FULL_NAME")
            )

            negosyo_payments = negosyo_payments.select(["ACCOUNT_NUMBER", "FULL_NAME", "CREATED_DATE", "TOTAL_PAYMENT", "TRANSACTION_ID", "PRODUCT_NAME"])

            negosyo_payments = negosyo_payments.with_columns(
                pl.col("ACCOUNT_NUMBER").cast(pl.Int64),
                pl.col("FULL_NAME").cast(pl.Utf8),
                pl.col("CREATED_DATE").cast(pl.Date),
                pl.col("TOTAL_PAYMENT").cast(pl.Float64),
                pl.col("TRANSACTION_ID").cast(pl.Utf8),
                pl.col("PRODUCT_NAME").cast(pl.Utf8)
            )
        else:
            negosyo_payments = None

        if sme_raw_file is not None:
            sme_payments = pl.read_excel(sme_raw_file)

            sme_payments = sme_payments.with_columns(
                (pl.col("AUTH_SIG_FIRST_NAME") + " " + pl.col("AUTH_SIG_LAST_NAME")).alias("FULL_NAME")
            )

            sme_payments = sme_payments.select(["ACCOUNT_NUMBER", "FULL_NAME", "CREATED_DATE", "TOTAL_PAYMENT", "TRANSACTION_ID", "PRODUCT_NAME"])

            sme_payments = sme_payments.with_columns(
                pl.col("ACCOUNT_NUMBER").cast(pl.Int64),
                pl.col("FULL_NAME").cast(pl.Utf8),
                pl.col("CREATED_DATE").cast(pl.Date),
                pl.col("TOTAL_PAYMENT").cast(pl.Float64),
                pl.col("TRANSACTION_ID").cast(pl.Utf8),
                pl.col("PRODUCT_NAME").cast(pl.Utf8)
            )
        else:
            sme_payments = None

        payment_files = [mayacredit_payments, negosyo_payments, sme_payments]

        payments_sheet = pl.DataFrame()
        
        for payments in payment_files:
            if payments is not None:
                payments_sheet = payments_sheet.vstack(payments)

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