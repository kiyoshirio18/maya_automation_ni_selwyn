import polars as pl
import streamlit as st
from io import BytesIO
from resources.excel_tools import save_xlsx, xlsx_to_xls

st.header("Payments Uploading")
st.write(f''':green[:green-background[PAYMENT TAG]] :green[:green-background[CH CODE]] :green[:green-background[ACC NUMBER]] :green[:green-background[PAYMENT DATE]] :green[:green-background[PAYMENT AMOUNT]]''')


if 'payment_autostatus' not in st.session_state:
    st.session_state.payment_autostatus = None

if 'volare_payments' not in st.session_state:
    st.session_state.volare_payments = None

if 'payments_raw_file' not in st.session_state:
    st.session_state.payments_raw_file = None

with st.form(key="payments"):
    file = st.file_uploader("Upload Payment File", type=["xlsx"])

    if st.form_submit_button("SUBMIT", use_container_width=True):
        st.session_state.payments_raw_file = pl.read_excel(file)

        #FOR VOLARE
        volare_payments = st.session_state.payments_raw_file.select(["ACC NUMBER", "PAYMENT AMOUNT", "PAYMENT DATE", "PAYMENT TAG"])

        volare_payments = volare_payments.with_columns(
            pl.lit("PAID").alias("STATUS")
        )

        st.session_state.volare_payments = volare_payments

        #FOR BCRM
        bcrm_payments = st.session_state.payments_raw_file.select(["PAYMENT TAG", "CH CODE", "PAYMENT DATE", "PAYMENT AMOUNT", "MAYA PRODUCT NAME"])

        bcrm_payments = bcrm_payments.with_columns(
            pl.when((pl.col("PAYMENT AMOUNT") == 15) & (pl.col("MAYA PRODUCT NAME") == "negosyoAdvance"))
            .then(pl.lit("NEGOSYO TRANSACTION FEE"))
            .otherwise(pl.lit("CLIENT PAID WITH BLASTING"))
            .alias("SUBSTATUS")
        )

        ptp_status = bcrm_payments.with_columns(
            pl.lit("PTP NEW").alias("STATUS"),
            pl.struct(["PAYMENT AMOUNT", "PAYMENT DATE"]).map_elements(
                lambda x: f"CLIENT WILL SETTLE PHP {x["PAYMENT AMOUNT"]:.2f} ON {x["PAYMENT DATE"].strftime("%m/%d/%Y")}",
                return_dtype=pl.Utf8
                ).alias("REMARKS"),
            (pl.col("PAYMENT DATE").cast(pl.Datetime) + pl.duration(hours = 8, minutes = 0)).alias("BARCODE DATE"),
            pl.col("PAYMENT DATE").alias("END DATE"),
            pl.lit(None).alias("OR NUMBER"),
            pl.lit(None).alias("NEW ADDRESS"),
            pl.lit(None).alias("NEW CONTACT")
        )

        ptp_status = ptp_status.rename({
            "PAYMENT TAG": "AGENT",
            "PAYMENT DATE": "START DATE"
        })

        ptp_status = ptp_status.select(["CH CODE", "STATUS", "SUBSTATUS", "PAYMENT AMOUNT", "START DATE", "END DATE", "OR NUMBER", "REMARKS", "NEW ADDRESS", "NEW CONTACT", "AGENT", "BARCODE DATE"])

        payment_status = bcrm_payments.with_columns(
            pl.lit("PAYMENT").alias("STATUS"),
            pl.struct(["PAYMENT AMOUNT", "PAYMENT DATE"]).map_elements(
                lambda x: f"CLIENT SETTLED PHP {x["PAYMENT AMOUNT"]:.2f} ON {x["PAYMENT DATE"].strftime("%m/%d/%Y")}",
                return_dtype=pl.Utf8
                ).alias("REMARKS"),
            (pl.col("PAYMENT DATE").cast(pl.Datetime) + pl.duration(hours = 8, minutes = 5)).alias("BARCODE DATE"),
            pl.col("PAYMENT DATE").alias("END DATE"),
            pl.lit(None).alias("OR NUMBER"),
            pl.lit(None).alias("NEW ADDRESS"),
            pl.lit(None).alias("NEW CONTACT")
        )
        payment_status = payment_status.rename({
            "PAYMENT TAG": "AGENT",
            "PAYMENT DATE": "START DATE"
        })

        payment_status = payment_status.select(["CH CODE", "STATUS", "SUBSTATUS", "PAYMENT AMOUNT", "START DATE", "END DATE", "OR NUMBER", "REMARKS", "NEW ADDRESS", "NEW CONTACT", "AGENT", "BARCODE DATE"])

        st.session_state.payment_autostatus = ptp_status.vstack(payment_status)
    
if st.session_state.payments_raw_file is not None:
    st.dataframe(st.session_state.payments_raw_file, use_container_width=True)

if st.session_state.payment_autostatus is not None:
    formatting = {"PAYMENT AMOUNT": "0.00"}

    date = st.session_state.payment_autostatus["END DATE"].max().strftime("%m%d%y")

    st.download_button(
        label = "Download Payments Autostatus",
        data = xlsx_to_xls(save_xlsx(st.session_state.payment_autostatus, formatting)),
        file_name = f"maya_payments_autostatus_{date}.xls",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if st.session_state.volare_payments is not None:
    formatting = {
        "ACC NUMBER": "0",
        "PAYMENT AMOUNT": "0.00",
        "PAYMENT DATE": "mm/dd/yyyy"
    }

    date = st.session_state.volare_payments["PAYMENT DATE"].max().strftime("%m%d%y")

    st.download_button(
        label = "Download Volare Payments",
        data = save_xlsx(st.session_state.volare_payments, formatting),
        file_name = f"maya_volare_payments_{date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )