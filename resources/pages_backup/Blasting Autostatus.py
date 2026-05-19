import polars as pl
import streamlit as st
from datetime import datetime
from io import BytesIO
from resources.excel_tools import save_xlsx, xlsx_to_xls

st.header("Blasting Autostatus")
st.write(f''':green[:green-background[CHCODE]] :green[:green-background[MOBILE PROPER]]''')

date = None
filename = None

if 'import_file' not in st.session_state:
    st.session_state.import_file = None

col1, col2 = st.columns(2)

with st.form(key="blasting_as"):
    file = st.file_uploader("Upload Blasting Import File", type=["xlsx"])
    col1, col2 = st.columns(2)
    option = col1.selectbox("Blasting", ("Voice", "SMS"))
    date = col2.date_input("Blasting Date", value="today")

    if option == "Voice":
        remark = "SENT VOICE BLAST TEMPLATE - MAYA V2"
        substatus = "VOICE MESSAGE PROMPT"
        filename = "maya_vb_autostat"

    if option == "SMS":
        remark = "SENT SMS TEMPLATE - ODREM3"
        substatus = "BLAST SMS"
        filename = "maya_sms_autostat"

    if st.form_submit_button("SUBMIT", use_container_width=True):
        if file is not None:
            import_file = pl.read_excel(file)
            import_file = import_file.select(["CHCODE", "MOBILE PROPER"])
        
        barcode_date = datetime.combine(date, datetime.min.time())
        barcode_date = barcode_date.replace(hour=10, minute=0, second=0)

        st.session_state.import_file = import_file.with_columns(
            pl.col("MOBILE PROPER").map_elements(lambda x: f"{x} {remark}", return_dtype=pl.Utf8).alias("REMARK"),
            pl.lit(barcode_date).alias("BARCODE_DATE"),
            pl.lit("LETTER SENT").alias("STATUS"),
            pl.lit(f"{substatus}").alias("SUBSTATUS"),
            pl.lit("MSPM").alias("AGENT"),
            pl.lit(None).alias("AMOUNT"),
            pl.lit(None).alias("START_DATE"),
            pl.lit(None).alias("END_DATE"),
            pl.lit(None).alias("OR_NUMBER"),
            pl.lit(None).alias("NEW_ADDRESS"),
            pl.lit(None).alias("NEW_CONTACT"),
        )

if st.session_state.import_file is not None:
    st.session_state.import_file = st.session_state.import_file.select(["CHCODE", "STATUS", "SUBSTATUS", "AMOUNT", "START_DATE", "END_DATE", "OR_NUMBER", "REMARK", "NEW_ADDRESS", "NEW_CONTACT", "AGENT", "BARCODE_DATE"])
    st.dataframe(st.session_state.import_file, use_container_width=True)

    st.download_button(
        label = "Download .xls",
        data = xlsx_to_xls(save_xlsx(st.session_state.import_file, formatting = None)),
        file_name = f"{filename}_{date.strftime("%m%d%Y")}.xls",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
