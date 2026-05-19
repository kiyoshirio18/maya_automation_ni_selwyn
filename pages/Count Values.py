import streamlit as st
import polars as pl
from io import BytesIO
import msoffcrypto

st.header("Count Values")
st.write(f''':green[:green-background[Masterfile]]''')

with st.form(key="blasting"):
    file = st.file_uploader("Upload XLSX", type=["xlsx"])

    passkey = st.text_input("Password", type="password")

    if st.form_submit_button("SUBMIT", use_container_width=True):
        if file is not None:
            excel_decrypted = BytesIO()
            with BytesIO(file.read()) as f:
                excel_file = msoffcrypto.OfficeFile(f)
                excel_file.load_key(password = passkey)
                excel_file.decrypt(excel_decrypted)
        
        df = pl.read_excel(excel_decrypted, sheet_name="ACTIVE", schema_overrides={
        'ID_NUMBER': pl.Utf8,  # Read as string to avoid conversion issues
        'ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE': pl.Utf8,  # Read as string to avoid conversion issues
    })
        col1, col2 = st.columns(2)

        col1.markdown("Count Per Product")
        col1.dataframe(df.group_by("PRODUCT_NAME").agg(
            pl.len().alias("Count"),
            pl.sum("OB"),
            pl.col("NAME").unique().len().alias("Unique Count")
        ).sort("PRODUCT_NAME"), use_container_width=True)
        col2.markdown("Total Accounts Per Placement")
        col2.dataframe(pl.Series(df.select("PLACEMENT")).value_counts().sort("PLACEMENT"), use_container_width=True)

        col1.markdown("Total Count per DPD Bucket")
        bucket = pl.DataFrame({"DPD BUCKET": ['1 - 30 DPD', '31 - 60 DPD', '61 - 90 DPD', '91 - 120 DPD', '121 - 150 DPD', '151 - 180 DPD', '181 DPD & UP']})
        email = df.filter(pl.col("EMAIL_ADDRESS").is_not_null()).group_by("DPD BUCKET").agg(
            pl.col("EMAIL_ADDRESS").len().alias("EMAIL COUNT")
        )
        bucket = bucket.join(email, on="DPD BUCKET", how="left")
        col2.dataframe(email)
        col1.dataframe(bucket.join(df.group_by("DPD BUCKET").agg(
            pl.col("DPD BUCKET").len().alias("COUNT"),
        ), on="DPD BUCKET", how="left"), use_container_width=True)
