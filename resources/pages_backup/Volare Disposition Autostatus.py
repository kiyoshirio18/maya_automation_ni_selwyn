import polars as pl
import streamlit as st
import json
from io import BytesIO
from datetime import datetime

st.header("Volare Disposition Autostatus")
st.write(f''':green[:green-background[Volare Daily Remark Report]]''')

if st.session_state.maya_active_pout is None:
    st.markdown(''':red[:red-background[Refresh Accounts on Main Page]]''')

with open('./resources/volare_status.json', 'r') as file:
    volare_status = json.load(file)

with open('./resources/volare_substatus.json', 'r') as file:
    volare_substatus = json.load(file)

with open('./resources/agent_code.json', 'r') as file:
    agent_code = json.load(file)

call_status = [
    "BULK SMS SENT",
    "BUSY",
    "DNC - HOLD EFFORT",
    "DROPPED",
    "JUNK - DECEASED CLIENT",
    "JUNK - NO ACTIVE NUMBER",
    "NEGATIVE - BUSY",
    "NEGATIVE - NO ANSWER",
    "NEGATIVE - VOICE CLIP",
    "NEGATIVE - WRONG NUMBER",
    "POSITIVE - 3RD PARTY CONTACTED",
    "POSITIVE - BUSY",
    "POSITIVE - FAILED ID",
    "POSITIVE - FIELD VISIT RESULT",
    "POSITIVE - HANG UP",
    "POSITIVE - SOCIAL MEDIA POSITIVE",
    "POSITIVE CONTACT - CALLBACK",
    "POSITIVE CONTACT - DISPUTE",
    "POSITIVE CONTACT - EMAIL RESPONSIVE",
    "POSITIVE CONTACT - RPC CALL DISCONNECTED",
    "POSITIVE CONTACT - RPC REFUSE TO PAY",
    "POSITIVE CONTACT - SMS RESPONSIVE",
    "POSITIVE CONTACT - UNDERNEGO",
    "POSITIVE CONTACT - VIBER RESPONSIVE",
    "PTP - RPC PTP FFUP",
    "RNA",
    "RTP",
    "UNTC",
    "VM"
]

ptp_status = [
    "PTP",
    "PTP - RPC PTP DISCOUNTED OTP",
    "PTP - RPC PTP FULL PAYMENT",
    "PTP - RPC PTP PARTIAL",
    "PTP - SPLIT PAYMENT"
]

maya_final = None
date = None

#Remove duplicate accounts and retain the most recent ACCOUNTS based on ENDO DATE
maya_accounts = st.session_state.maya_active_pout
maya_accounts = maya_accounts.sort(by=["ACCOUNT NUM", "ENDO DATE"], descending=[False, True])
maya_accounts = maya_accounts.unique(subset=["ACCOUNT NUM"], keep="first")

def save_excel(_df: pl.DataFrame) -> bytes:
    output = BytesIO()
    if _df is not None:
        _df.write_excel(
            output, 
            autofit=True)
    output.seek(0)
    return output.getvalue()

with st.form(key="volare_status"):
    file = st.file_uploader("Upload Volare Daily Remark", type=["xlsx"])

    if st.form_submit_button("SUBMIT", use_container_width=True):
        if file is not None:
            maya_status = pl.read_excel(file, schema_overrides={
                "Date": pl.Date,
                "PTP Date": pl.Date
            })
        
        maya_call_status = maya_status.filter(pl.col("Status").is_in(call_status))
        maya_ptp_status = maya_status.filter(pl.col("Status").is_in(ptp_status))

        maya_ptp_status = maya_ptp_status.select(["Time","Account No.", "Status","Remark","Remark By","PTP Amount","PTP Date"])
        maya_call_status = maya_call_status.select(["Time","Account No.", "Status","Remark","Remark By"])

        maya_ptp_status = maya_ptp_status.with_columns(
            pl.col("Status").map_elements(lambda x : volare_status.get(x, None), return_dtype=pl.Utf8).alias("STATUS"),
            pl.col("Status").map_elements(lambda x : volare_substatus.get(x, None), return_dtype=pl.Utf8).alias("SUBSTATUS"),
            pl.col("Remark By").map_elements(lambda x : agent_code.get(x, None), return_dtype=pl.Utf8).alias("AGENT"),
            pl.col("PTP Amount").alias("AMOUNT"),
            pl.col("PTP Date").alias("START_DATE"),
            pl.col("PTP Date").alias("END_DATE"),
            pl.col("Account No.").cast(pl.Utf8),
            pl.lit(None).alias("OR_NUMBER"),
            pl.lit(None).alias("NEW_ADDRESS"),
            pl.lit(None).alias("NEW_CONTACT")
        )

        maya_call_status = maya_call_status.with_columns(
            pl.col("Status").map_elements(lambda x : volare_status.get(x, None), return_dtype=pl.Utf8).alias("STATUS"),
            pl.col("Status").map_elements(lambda x : volare_substatus.get(x, None), return_dtype=pl.Utf8).alias("SUBSTATUS"),
            pl.col("Remark By").map_elements(lambda x : agent_code.get(x, None), return_dtype=pl.Utf8).alias("AGENT"),
            pl.col("Account No.").cast(pl.Utf8),
            pl.lit(None).alias("START_DATE"),
            pl.lit(None).alias("END_DATE"),
            pl.lit(None).alias("OR_NUMBER"),
            pl.lit(None).alias("NEW_ADDRESS"),
            pl.lit(None).alias("NEW_CONTACT"),
            pl.lit(None).alias("AMOUNT")
        )

        maya_call_status = maya_call_status.with_columns(
            pl.col("AGENT").fill_null("MSPM")
        )

        maya_ptp_status = maya_ptp_status.with_columns(
            pl.col("AGENT").fill_null("MSPM")
        )

        maya_ptp_status = maya_ptp_status.join(maya_accounts.select(["ACCOUNT NUM", "CHCODE"]), left_on="Account No.", right_on="ACCOUNT NUM", how="left")
        maya_call_status = maya_call_status.join(maya_accounts.select(["ACCOUNT NUM", "CHCODE"]), left_on="Account No.", right_on="ACCOUNT NUM", how="left")

        maya_ptp_status = maya_ptp_status.select(["CHCODE", "STATUS", "SUBSTATUS", "AMOUNT", "START_DATE", "END_DATE", "OR_NUMBER", "Remark", "NEW_ADDRESS", "NEW_CONTACT", "AGENT", "Time"])
        maya_call_status = maya_call_status.select(["CHCODE", "STATUS", "SUBSTATUS", "AMOUNT", "START_DATE", "END_DATE", "OR_NUMBER", "Remark", "NEW_ADDRESS", "NEW_CONTACT", "AGENT", "Time"])

        maya_ptp_status = maya_ptp_status.filter(pl.col("CHCODE").is_not_null() & pl.col("AGENT").is_not_null())
        maya_ptp_status = maya_ptp_status.filter(~((pl.col("STATUS") == "PTP NEW") & (pl.col("AMOUNT") == 0)))
        maya_call_status = maya_call_status.filter(pl.col("CHCODE").is_not_null() & pl.col("AGENT").is_not_null())

        maya_final = pl.concat([maya_ptp_status, maya_call_status], how="vertical")
        date = (maya_final["Time"].max()).strftime("%m%d%y")

        st.dataframe(maya_final)

if maya_final is not None:
    
    st.download_button(
        label = "Download .xlsx",
        data = save_excel(maya_final),
        file_name = f"maya_volare_autostatus_{date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )