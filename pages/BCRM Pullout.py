import streamlit as st
import polars as pl
from datetime import datetime
from io import BytesIO
from resources.excel_tools import save_xlsx, xlsx_to_xls

st.header("Pullout")
st.write(f''':green[:green-background[CHCODE]]''')

if "pout_autostat" not in st.session_state:
    st.session_state.pout_autostat = None

if "pout_reshuffle" not in st.session_state:
    st.session_state.pout_reshuffle = None

file_name = None
pout_date = None

col4, col5 = st.columns(2)

def pullout(chcode, option, pout_date):
    chcodes = chcode.strip().split("\n")
    pout_chcodes = pl.DataFrame({"CHCODE": chcodes})
    st.session_state.pout_autostat = pout_chcodes
    st.session_state.pout_reshuffle = pout_chcodes

    st.session_state.pout_reshuffle = st.session_state.pout_reshuffle.with_columns(
        pl.lit("POUT").alias("TAGGING")
    )

    if option == "Recall":
        substatus = "RETURNED TO BANK"
        remark = f"RECALL BY BANK {pout_date}"
        file_name = "maya_recall"
    
    if option == "Fully Paid":
        substatus = "FULLY PAID WITH CLOSURE REQUEST"
        remark = f"PULLOUT FULLY PAID {pout_date}"
        file_name = "maya_fullypaid"

    barcode_date = datetime.combine(pout_date, datetime.min.time())
    barcode_date = barcode_date.replace(hour=22, minute=0, second=0)

    st.session_state.pout_autostat = st.session_state.pout_autostat.with_columns(
        pl.lit("PULLED OUT").alias("STATUS"),
        pl.lit(substatus).alias("SUBSTATUS"),
        pl.lit(None).alias("AMOUNT"),
        pl.lit(None).alias("START_DATE"),
        pl.lit(None).alias("END_DATE"),
        pl.lit(None).alias("OR_NUMBER"),
        pl.lit(remark).alias("REMARK"),
        pl.lit(None).alias("NEW_ADDRESS"),
        pl.lit(None).alias("NEW_CONTACT"),
        pl.lit("POUT").alias("AGENT"),
        pl.lit(barcode_date).alias("BARCODE_DATE")
    )

with col4.form(key="pullout"):
    chcode = st.text_area("Input CHCODES", height=500)

    col1, col2, col3= st.columns(3, vertical_alignment="bottom")
    option = col1.selectbox(
        "Option",
        ("Fully Paid", "Recall")
    )

    if option == "Recall":
        file_name = "maya_recall"
    
    if option == "Fully Paid":
        file_name = "maya_fullypaid"

    pout_date = col2.date_input("POUT Date", value="today")
 
    if col3.form_submit_button(use_container_width=True):
        pullout(chcode, option, pout_date)

if st.session_state.pout_autostat is not None and st.session_state.pout_reshuffle is not None:
    col5.dataframe(st.session_state.pout_autostat, height=525)
    col5.download_button(
        label = "Download Reshuffle File",
        data = xlsx_to_xls(save_xlsx(st.session_state.pout_reshuffle, None)),
        file_name = f"{file_name}_reshuffle_{pout_date}.xls",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    col5.download_button(
        label = "Download Autostatus File",
        data = xlsx_to_xls(save_xlsx(st.session_state.pout_autostat, None)),
        file_name = f"{file_name}_autostatus_{pout_date}.xls",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


