import polars as pl
import streamlit as st
import msoffcrypto
import io
import pickle

st.header("Digital Report")

def decrypt_excel(input_data: bytes | io.BytesIO, password: str) -> bytes:
    """
    Decrypts an Excel file from bytes input and returns decrypted bytes output.
    
    :param input_bytes: Encrypted Excel file in bytes
    :param password: Password to decrypt the file
    :return: Decrypted Excel file in bytes
    """
    if isinstance(input_data, io.BytesIO):
        input_stream = input_data  # Use as is if already a BytesIO object
    else:
        input_stream = io.BytesIO(input_data)  # Convert bytes to BytesIO
    output_stream = io.BytesIO()
    
    # Open the encrypted file
    file = msoffcrypto.OfficeFile(input_stream)
    file.load_key(password=password)
    file.decrypt(output_stream)
    
    return output_stream.getvalue()

with st.form(key="digital_report"):
    col1, col2 = st.columns(2, vertical_alignment="bottom")
    email_blast_file = col2.file_uploader("Upload Email Blasting File", type="xlsx")
    sms_blast_file = col2.file_uploader("Upload SMS Blasting File", type="xlsx")
    daily_remark_file = col1.file_uploader("Daily Remark File", type="xlsx")
    masterfile_encrypted = col1.file_uploader("Upload Masterfile", type="xlsx")
    password = st.text_input("Masterfile Password", type="password")

    if st.form_submit_button(use_container_width=True):
        masterfile = pl.read_excel(decrypt_excel(masterfile_encrypted, "Maya@2025"), schema_overrides={"ACCOUNT NUMBER": pl.Int64})
        masterfile = masterfile.select(["ACCOUNT NUMBER", "DPD BUCKET"])

        email_blast = pl.read_excel(email_blast_file)
        email_blast = email_blast.filter(pl.col("ACCOUNT_NUMBER").is_not_null())
        email_blast_accs = email_blast["ACCOUNT_NUMBER"].unique().to_list()

        sms_blast = pl.read_excel(sms_blast_file)

        daily_remark = pl.read_excel(daily_remark_file, schema_overrides={"Account No.": pl.Int64})
        connected_accs = daily_remark.filter(pl.col("Call Status") == "CONNECTED")["Account No."].unique().to_list()

        unconnected_accs = masterfile.filter(~(pl.col("ACCOUNT NUMBER").is_in(connected_accs)))
        email_accs = masterfile.filter(pl.col("ACCOUNT NUMBER").is_in(email_blast_accs))
        sms_accs = masterfile.filter(pl.col("ACCOUNT NUMBER").is_in(sms_blast["ACCOUNT_NUMBER"]))

        data_list = {
            "unconnected": {
                "data": unconnected_accs
            },
            "email": {
                "data": email_accs
            },
            "sms": {
                "data": sms_accs
            }
        }

        for keys, values in data_list.items():
            values["summary"] = values["data"].group_by("DPD BUCKET").agg(
                pl.col("ACCOUNT NUMBER").len().alias(f"{keys}_count")
            )

        summary = pl.DataFrame({"bucket": ["1 - 30 DPD", "31 - 60 DPD", "61 - 90 DPD", "91 - 120 DPD", "121 - 150 DPD", "151 - 180 DPD", "181 DPD & UP"]})

        for values in data_list.values():
            summary = summary.join(values["summary"], left_on="bucket", right_on="DPD BUCKET", how="left")

        st.dataframe(summary, use_container_width=True)