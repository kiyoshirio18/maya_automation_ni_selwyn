import streamlit as st
import polars as pl
import msoffcrypto
import json
from io import BytesIO

st.header("Claiming Paid Autostat")
st.write(f''':green[:green-background[ACC NUMBER]] :green[:green-background[ACC NAME]] :green[:green-background[PAYMENT DATE]] :green[:green-background[PAYMENT AMOUNT]] :green[:green-background[TRANSACTION ID]]''')

if "import_remarks_file" not in st.session_state:
    st.session_state.import_remarks_file = None

if "import_remarks_file_name" not in st.session_state:
    st.session_state.import_remarks_file_name = None

with st.form(key="claiming_paid"):
    file = st.file_uploader("Upload Untagged Payments", type=["xlsx"])

    if st.form_submit_button("SUBMIT", use_container_width=True):
        untagged_payments = pl.read_excel(file, schema_overrides={
            "ACC NUMBER": pl.Int64,
            "ACC NAME": pl.Utf8,
            "PAYMENT DATE": pl.Datetime,
            "PAYMENT AMOUNT": pl.Float64,
            "TRANSACTION ID": pl.Utf8
        })

        untagged_payments = untagged_payments.rename({
            "ACC NUMBER": "Account Number",
            "PAYMENT DATE": "Remark Date",
            "PAYMENT AMOUNT": "Claim Paid Amount"
        })

        import_remarks = untagged_payments.with_columns(
            pl.lit("PAYMENT - PAID").alias("Action Status"),
            (pl.col("Account Number").cast(str)
            + " PAID PHP"
            + pl.col("Claim Paid Amount").map_elements(lambda x: f"{x:.2f}", return_dtype=pl.Utf8)
            + " ON "
            + pl.col("Remark Date").dt.strftime("%Y-%m-%d")
            + " ID:"
            + pl.col("TRANSACTION ID").cast(str)
            ).alias("Remark"),
            pl.col("Remark Date").cast(pl.Date).alias("Claim Paid Date"),
            pl.lit("SPMADRID").alias("Remark By"),
            pl.lit(None).alias("PTP Date"),
            pl.lit(None).alias("Reason For Default"),
            pl.lit(None).alias("Field Visit Date"),
            pl.lit(None).alias("Next Call Date"),
            pl.lit(None).alias("PTP Amount"),
            pl.lit(None).alias("Phone No"),
            pl.lit(None).alias("Relation"),
        )

        import_remarks = import_remarks.select(["Account Number", "Action Status", "Remark Date", "PTP Date", "Reason For Default",
                                                "Field Visit Date", "Remark", "Next Call Date", "PTP Amount", "Claim Paid Amount",
                                                "Remark By", "Phone No", "Relation", "Claim Paid Date"])
        
        st.dataframe(import_remarks)

        st.session_state.import_remarks_file = BytesIO()
        st.session_state.import_remarks_file_name = f"maya_claiming_paid_{import_remarks["Claim Paid Date"].min().strftime("%m%d%Y")}-{import_remarks["Claim Paid Date"].max().strftime("%m%d%Y")}.xlsx"

        import_remarks.write_excel(
            st.session_state.import_remarks_file,
            dtype_formats={
                pl.Int64: "0",
                pl.Float64: "0.00",
                pl.Date: "mm/dd/yyyy",
                pl.Datetime: "mm/dd/yyyy hh:mm:ss"
            }
        )

if st.session_state.import_remarks_file is not None:
    st.session_state.import_remarks_file.seek(0)
    st.download_button(
        label = "claiming paid autostat",
        data = st.session_state.import_remarks_file.getvalue(),
        file_name = st.session_state.import_remarks_file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
