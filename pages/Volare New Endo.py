import streamlit as st
import polars as pl
import msoffcrypto
import json
from io import BytesIO

st.header("Volare New Endo Uploading")
st.write(f''':green[:green-background[Masterfile]] :green[:green-background[All Endo]]''')

if "mc121" not in st.session_state:
    st.session_state.mc121 = None

if "mc181" not in st.session_state:
    st.session_state.mc181 = None

if "negosyo_sme" not in st.session_state:
    st.session_state.negosyo_sme = None

if "file_date" not in st.session_state:
    st.session_state.file_date = None

if "instacash" not in st.session_state:
    st.session_state.instacash = None

def mobile_number(x):

    if x.startswith("63"):
        return "0" + x[2:]
    
    if x.startswith("9") and len(x) == 10:
        return "0" + x
    
    return x

def save_excel(_df: pl.DataFrame) -> bytes:
    output = BytesIO()
    if _df is not None:
        _df.write_excel(
            output, 
            autofit=True,
            column_formats={
            "ACCOUNT_NUMBER": "0",
            "OB": "0.00",
            "RECEIVED DATE": "mm/dd/yyyy",
            "BIRTH_DATE": "mm/dd/yyyy"
        })
    output.seek(0)
    return output.getvalue()

with open('./resources/maya_volare_cycle.json', 'r') as file:
    volare_cycle = json.load(file)

with open('./resources/agent_code_bcrm_volare.json', 'r') as file:
    volare_agents = json.load(file)

with st.form(key="blasting"):
    file = st.file_uploader("Upload Password Protected XLSX", type=["xlsx"])
    passkey = st.text_input("Password", type="password")

    if st.form_submit_button("SUBMIT", use_container_width=True):
        if file is not None:
            try:
                excel_decrypted = "./temp/decrypted.xlsx"
                with BytesIO(file.read()) as f:
                    excel_file = msoffcrypto.OfficeFile(f)
                    excel_file.load_key(password = passkey)
                    with open(excel_decrypted, 'wb') as df:
                        excel_file.decrypt(df)
                masterfile_df = pl.read_excel(excel_decrypted, sheet_name='ACTIVE', schema_overrides={
                    'ID_NUMBER': pl.Utf8,
                    'ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE': pl.Utf8,
                    'RECEIVED DATE': pl.Date,
                    'BIRTH_DATE': pl.Date,
                    'ACCOUNT NUMBER': pl.Int64
                })
            except msoffcrypto.exceptions.DecryptionError as e:
                if str(e) == "Unencrypted document" or str(e) == "No key specified":
                    print("Caught DecryptionError with message: Unencrypted document")
                    masterfile_df = pl.read_excel(file, schema_overrides={
                        'ID_NUMBER': pl.Utf8,
                        'ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE': pl.Utf8,
                        'RECEIVED DATE': pl.Date,
                        'BIRTH_DATE': pl.Date,
                        'ACCOUNT NUMBER': pl.Int64
                    })
                else:
                    print(f"Caught DecryptionError with a different message: {e}")


        masterfile_df = masterfile_df.filter(pl.col("ENDO STAT") == "NEW ENDO")
        masterfile_df = masterfile_df.select(["PLACEMENT", "ACCOUNT_NUMBER", "CHCODE", "DPD_", "TAGGING", "MOBILE PROPER", "OB", "RECEIVED DATE", "NAME", "BIRTH_DATE", "ALTERNATIVE_NUMBER", "EMAIL_ADDRESS", "PRESENT_ADDRESS", "NATURE_OF_WORK", "CONTACT_REFERENCE_MOBILE_NUMBER", "CONTACT_REFERENCE_FIRST_NAME", "CONTACT_REFERENCE_LAST_NAME", "CONTACT_REFERENCE_RELATIONSHIP", "TU_NUMBER_1", "TU_NUMBER_2", "TU_NUMBER_3", "TU_NUMBER_4", "TU_NUMBER_5"])
        st.session_state.file_date = masterfile_df["RECEIVED DATE"].max()

        masterfile_df = masterfile_df.with_columns(
            pl.col("RECEIVED DATE").dt.date().alias("RECEIVED DATE"),
            pl.col("BIRTH_DATE").dt.date().alias("BIRTH_DATE"),
            #pl.lit("SPMADRID").alias("TAGGING"),
            #pl.col("TAGGING").map_elements(lambda x: volare_agents.get(x, None), return_dtype=pl.Utf8).alias("TAGGING"),
            pl.col("PLACEMENT").map_elements(lambda x: volare_cycle.get(x, None), return_dtype=pl.Utf8).alias("CYCLE"),
            pl.col("ALTERNATIVE_NUMBER").cast(pl.Utf8),
            pl.col("CONTACT_REFERENCE_MOBILE_NUMBER").cast(pl.Utf8),
            (pl.col("CONTACT_REFERENCE_FIRST_NAME") + " " + pl.col("CONTACT_REFERENCE_LAST_NAME")).alias("CONTACT_REFERENCE_NAME")
        )

        masterfile_df = masterfile_df.with_columns(
            pl.col("ALTERNATIVE_NUMBER").map_elements(mobile_number, return_dtype=pl.Utf8).alias("ALTERNATIVE_NUMBER"),
            pl.col("CONTACT_REFERENCE_MOBILE_NUMBER").map_elements(mobile_number, return_dtype=pl.Utf8).alias("CONTACT_REFERENCE_MOBILE_NUMBER"),
            pl.col("TU_NUMBER_1").map_elements(mobile_number, return_dtype=pl.Utf8).alias("TU_NUMBER_1"),
            pl.col("TU_NUMBER_2").map_elements(mobile_number, return_dtype=pl.Utf8).alias("TU_NUMBER_2"),
            pl.col("TU_NUMBER_3").map_elements(mobile_number, return_dtype=pl.Utf8).alias("TU_NUMBER_3"),
            pl.col("TU_NUMBER_4").map_elements(mobile_number, return_dtype=pl.Utf8).alias("TU_NUMBER_4"),
            pl.col("TU_NUMBER_5").map_elements(mobile_number, return_dtype=pl.Utf8).alias("TU_NUMBER_5")
        )

        masterfile_df = masterfile_df.with_columns(
            pl.col("ALTERNATIVE_NUMBER").map_elements(lambda x: None if x == "0" else x, return_dtype=pl.Utf8).alias("ALTERNATIVE_NUMBER"),
            pl.col("TU_NUMBER_1").map_elements(lambda x: None if x == "0" else x, return_dtype=pl.Utf8).alias("TU_NUMBER_1"),
            pl.col("TU_NUMBER_2").map_elements(lambda x: None if x == "0" else x, return_dtype=pl.Utf8).alias("TU_NUMBER_2"),
            pl.col("TU_NUMBER_3").map_elements(lambda x: None if x == "0" else x, return_dtype=pl.Utf8).alias("TU_NUMBER_3"),
            pl.col("TU_NUMBER_4").map_elements(lambda x: None if x == "0" else x, return_dtype=pl.Utf8).alias("TU_NUMBER_4"),
            pl.col("TU_NUMBER_5").map_elements(lambda x: None if x == "0" else x, return_dtype=pl.Utf8).alias("TU_NUMBER_5")
        )

        masterfile_df = masterfile_df.select(["CYCLE", "DPD_", "CHCODE", "TAGGING", "ACCOUNT_NUMBER", "NAME", "OB", "RECEIVED DATE", "MOBILE PROPER", "ALTERNATIVE_NUMBER", "CONTACT_REFERENCE_MOBILE_NUMBER", "CONTACT_REFERENCE_NAME", "CONTACT_REFERENCE_RELATIONSHIP", "BIRTH_DATE", "PRESENT_ADDRESS", "EMAIL_ADDRESS","NATURE_OF_WORK", "TU_NUMBER_1", "TU_NUMBER_2", "TU_NUMBER_3", "TU_NUMBER_4", "TU_NUMBER_5"])

        masterfile_df = masterfile_df.rename({
            "DPD_": "DPD"
        })

        masterfile_df = masterfile_df.with_columns(
            [pl.col(col).str.replace("'", "") for col in masterfile_df.schema if masterfile_df.schema[col] == pl.Utf8]
        )
        
        # masterfile_df = masterfile_df.rename({
        #     "CYCLE": "PLACEMENT",
        #     "ALTERNATIVE": "ALTERNATIVE_NUMBER",
        #     "CONTACT_REFERENCE": "CONTACT_REFERENCE_MOBILE_NUMBER",
        #     "BIRTH_DATE": "BIRTH DATE",
        #     "PRESENT_ADDRESS": "ADDRESS"
        # })

        st.session_state.mc121 = masterfile_df.filter(
            pl.col("CYCLE") == "MC 121-150 DPD"
        )

        st.session_state.mc181 = masterfile_df.filter(
            (pl.col("CYCLE") == "MC 181DPD UP") |
            (pl.col("CYCLE") == "Maya PayLater 181DPD")
        )

        st.session_state.negosyo_sme = masterfile_df.filter(
            pl.col("CYCLE").str.starts_with("Maya Negosyo") | 
            pl.col("CYCLE").str.starts_with("Maya SME")
        )

        st.session_state.instacash = masterfile_df.filter(
            pl.col("CYCLE").str.starts_with("Maya ITC")
        )

if st.session_state.mc121 is not None and st.session_state.mc121.height > 0:
    st.download_button(
        label = "Download MC121 NEW ENDO",
        data = save_excel(st.session_state.mc121),
        file_name = f"maya_volare_upload_121DPD_{st.session_state.file_date.strftime("%m%d")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if st.session_state.mc181 is not None and st.session_state.mc181.height > 0:
    st.download_button(
        label = "Download MC181 NEW ENDO",
        data = save_excel(st.session_state.mc181),
        file_name = f"maya_volare_upload_181DPD_{st.session_state.file_date.strftime("%m%d")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if st.session_state.negosyo_sme is not None and st.session_state.negosyo_sme.height > 0:
    st.download_button(
        label = "Download Maya Negosyo SME NEW ENDO",
        data = save_excel(st.session_state.negosyo_sme),
        file_name = f"maya_volare_upload_negosyo_sme_{st.session_state.file_date.strftime("%m%d")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if st.session_state.instacash is not None and st.session_state.instacash.height > 0:
    st.download_button(
        label = "Download Maya Instacash NEW ENDO",
        data = save_excel(st.session_state.instacash),
        file_name = f"maya_volare_upload_instacash_{st.session_state.file_date.strftime("%m%d")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )