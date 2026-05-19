import polars as pl
import streamlit as st
import json
import msoffcrypto
import pickle
import os
from io import BytesIO
from datetime import datetime
from resources.excel_tools import cast_columns, save_xlsx
import resources.schema as schema

st.header("MCC Endorsement")
st.write(f''':green[:green-background[Endorsement Raw Files]] :green[:green-background[Masterfile]]''')

def get_dpd_bucket(days_past_due):
    """
    Takes an integer input representing days past due (DPD) and 
    returns the corresponding DPD bucket.

    Args:
        days_past_due (int): The number of days past due.

    Returns:
        str: The corresponding DPD bucket.
    """
    if days_past_due < 1:
        return "0 DPD"
    elif 1 <= days_past_due <= 30:
        return "1 - 30 DPD"
    elif 31 <= days_past_due <= 60:
        return "31 - 60 DPD"
    elif 61 <= days_past_due <= 90:
        return "61 - 90 DPD"
    elif 91 <= days_past_due <= 120:
        return "91 - 120 DPD"
    elif 121 <= days_past_due <= 150:
        return "121 - 150 DPD"
    elif 151 <= days_past_due <= 180:
        return "151 - 180 DPD"
    else:
        return "181 DPD & UP"
    
def mobile_number(x: str):

    if x.startswith("63"):
        return "0" + x[2:]
    
    if x.startswith("9") and len(x) == 10:
        return "0" + x
    
    return x

def check_date_status(input_date):
    # Get the current date
    current_date = datetime.now()

    # Compare the year and month of the input date with the current date
    if input_date.year == current_date.year and input_date.month == current_date.month:
        return "FRESH"
    else:
        return "SPILLOVER"

class TagAssigner:
    def __init__(self, index_file="./resources/maya_product_taggings.json"):
        self.index_file = index_file
        self.tag_index = self._load_index()

    def _load_index(self):
        if os.path.exists(self.index_file):
            try:
                with open(self.index_file, "r") as f:
                    return json.load(f)
            except json.JSONDecodeError:
                return {}
        return {}

    def save_index(self):
        with open(self.index_file, "w") as f:
            json.dump(self.tag_index, f)

    def assign_taggings(self, placement: str) -> str:
        if placement == "Maya Credit 121 - 150 DPD":
            return "PKDV"
        
        if placement == "Maya Credit 181 DPD & UP":
            tag_list = ["PKDV", "PJGF", "PKGV", "PVEM", "PAOB", "PNPL", "PAHD", "PRJL", "PJDX"]
            current_index = self.tag_index.get("Maya Credit 181 DPD & UP", 0)
            tag = tag_list[current_index]
            self.tag_index["Maya Credit 181 DPD & UP"] = (current_index + 1) % len(tag_list)
            return tag
        
        if (placement.startswith("Maya Negosyo Advance") or 
            placement.startswith("Maya Instacash") or 
            placement.startswith("Maya SME Flexi Loan")):
            tag_list = ["PZLB", "PSRR"]
            current_index = self.tag_index.get("Maya Negosyo SME Instacash", 0)
            tag = tag_list[current_index]
            self.tag_index["Maya Negosyo SME Instacash"] = (current_index + 1) % len(tag_list)
            return tag

        return None

with open("./resources/maya_mcc_mappings.json", "r") as file:
    maya_mcc_mapping = json.load(file)

with open("./resources/maya_product_placement.json", "r") as file:
    maya_product_placement = json.load(file)

masterfile_datatypes = schema.mcc_schema()

with open("./resources/maya_masterfile_formatting.json", "r") as file:
    masterfile_formatting = json.load(file)

with open("./resources/maya_product_taggings.json", "r") as file:
    maya_product_taggings = json.load(file)

if "masterfile_excel_file" not in st.session_state:
    st.session_state.masterfile_excel_file = None

if "new_endo_excel_file" not in st.session_state:
    st.session_state.new_endo_excel_file = None

if "masterfile" not in st.session_state:
    st.session_state.masterfile = None

if "new_endo" not in st.session_state:
    st.session_state.new_endo = None

if "bcrm_upload" not in st.session_state:
    st.session_state.bcrm_upload = None

if "bcrm_upload_file" not in st.session_state:
    st.session_state.bcrm_upload_file = None

if "masterfile_updated" not in st.session_state:
    st.session_state.masterfile_updated = None

if "masterfile_updated_file" not in st.session_state:
    st.session_state.masterfile_updated_file = None

if "masterfile_yesterday" not in st.session_state:
    st.session_state.masterfile_yesterday = None

if "notinendo" not in st.session_state:
    st.session_state.notinendo = None

if "endo_date" not in st.session_state:
    st.session_state.endo_date = None

with st.form(key="endorsement"):
    col1, col2 = st.columns(2)

    mcc = col1.file_uploader("Maya CC Endo Raw File", type="xlsx")

    masterfile = col2.file_uploader("Yesterday's Masterfile", type="xlsx")
    st.session_state.endo_date = col1.date_input("Endorsement Date", value="today")
    passkey = col2.text_input("Masterfile Password", type="password")

    if st.form_submit_button(use_container_width=True):
        with st.status(label="Processing...", expanded=True) as status:
            st.write("Decrypting Masterfile...")
            if masterfile is not None:
                try:
                    excel_decrypted = BytesIO()
                    with BytesIO(masterfile.read()) as f:
                        excel_file = msoffcrypto.OfficeFile(f)
                        excel_file.load_key(password = passkey)
                        excel_file.decrypt(excel_decrypted)
                    masterfile_yesterday = pl.read_excel(excel_decrypted, sheet_name="ACTIVE")
                    masterfile_yesterday = cast_columns(masterfile_yesterday, masterfile_datatypes)
                    masterfile_yesterday = masterfile_yesterday.with_columns(
                        pl.lit("OLD ENDO").alias("ENDO STATUS")
                    )
                    
                    st.session_state.masterfile_yesterday = masterfile_yesterday
                except msoffcrypto.exceptions.DecryptionError as e:
                    if str(e) == "Unencrypted document" or str(e) == "No key specified":
                        print("Caught DecryptionError with message: Unencrypted document")
                        masterfile_yesterday = pl.read_excel(masterfile, engine="openpyxl")
                        st.session_state.masterfile_yesterday = masterfile_yesterday
                    else:
                        print(f"Caught DecryptionError with a different message: {e}")

                st.write("Aligning Raw File Columns to Mastefile Headers...")
                files = {
                    'mcc_raw_file': mcc
                }

                for var_name, file in files.items():
                    if file is None:
                        globals()[var_name] = None
                    else:
                        globals()[var_name] = pl.read_excel(file, engine="openpyxl")
                
                # Dictionary to map raw files to their respective mappings
                raw_files = {
                    "mcc": {"file": globals().get("mcc_raw_file"), "mapping": maya_mcc_mapping},
                }

                # Loop through each raw file
                for file_name, data in raw_files.items():
                    raw_file = data["file"]
                    mapping = data["mapping"]
                    
                    # Check if the raw file exists (i.e., is not None)
                    if raw_file is not None:
                        # Aligning columns to Masterfile
                        masterfile = pl.DataFrame({
                            new_col: raw_file[existing_col] if existing_col in raw_file.columns else pl.Series(new_col, [None] * raw_file.height)
                            for new_col, existing_col in mapping.items()
                        })

                        # Cast columns to appropriate data types
                        masterfile = cast_columns(masterfile, masterfile_datatypes)

                        # Store or use the masterfile (negosyo_masterfile, credit_masterfile, sme_masterfile)
                        if file_name == "mcc":
                            mcc_masterfile = masterfile
                    else:
                        # Set the corresponding masterfile to None if raw_file is None
                        if file_name == "mcc":
                            mcc_masterfile = None
            
            st.write("Stacking Aligned Endo Files...")
            # List of masterfile variables
            masterfile = mcc_masterfile

            st.write("Filling in Necessary Columns...")
            #Filling in necessary columns
            masterfile = masterfile.with_columns(
                (pl.col("FIRST_NAME") + " " + pl.col("LAST_NAME")).str.to_uppercase().alias("FULL NAME"),
                (pl.col("CONTACT_REFERENCE_FIRST_NAME") + " " + pl.col("CONTACT_REFERENCE_LAST_NAME")).str.to_uppercase().alias("CONTACT_REFERENCE_FULL_NAME"),
                pl.col("DPD").map_elements(get_dpd_bucket, return_dtype=pl.Utf8).alias("DPD BUCKET"),
                pl.col("MOBILE_NUMBER_1").map_elements(mobile_number, return_dtype=pl.Utf8).alias("MOBILE PROPER"),
                pl.when(pl.col("ACCOUNT NUMBER").is_in(masterfile_yesterday["ACCOUNT NUMBER"])).then(pl.lit("OLD ENDO")).otherwise(pl.lit("NEW ENDO")).alias("ENDO STATUS")
            )

            masterfile = masterfile.with_columns(
                pl.col(f"tu_contact_number_{x}").map_elements(mobile_number, return_dtype=pl.Utf8).alias(f"tu_contact_number_{x}") for x in range(1,6)
            )

            masterfile = masterfile.with_columns(
                (pl.col("PRODUCT_NAME").replace_strict(maya_product_placement, default=None) + " " + pl.col("DPD BUCKET")).alias("PLACEMENT")
            )


            # Perform the join for each column
            masterfile = masterfile.join(
                masterfile_yesterday.select(["ACCOUNT NUMBER", "RECEIVED DATE"]), 
                on="ACCOUNT NUMBER", 
                how="left", 
                suffix="_new"  # No need for .lower()
            )

            # Update the column with the new values if they exist in the joined DataFrame
            masterfile = masterfile.with_columns(
                pl.when(pl.col(f"RECEIVED DATE_new").is_not_null())
                .then(pl.col(f"RECEIVED DATE_new"))
                .otherwise(pl.col("RECEIVED DATE"))
                .alias("RECEIVED DATE")
            )

            masterfile = masterfile.with_columns(
                pl.when(
                    pl.col("ENDO STATUS") == "NEW ENDO"
                    ).then(
                        pl.lit(st.session_state.endo_date).alias("RECEIVED DATE")
                    ).otherwise(
                        pl.col("RECEIVED DATE")
                    )
            )

            masterfile = masterfile.with_columns(
                pl.col("RECEIVED DATE").map_elements(check_date_status, return_dtype=pl.Utf8).alias("FRESH/SPILLOVER")
            )

            masterfile = masterfile.drop(["RECEIVED DATE_new"])

            st.session_state.masterfile = masterfile
            st.session_state.notinendo = st.session_state.masterfile_yesterday.join(st.session_state.masterfile, on="ACCOUNT NUMBER", how="anti")
            status.update(label="Done", state="complete", expanded=False)


if st.session_state.masterfile is not None:
    st.dataframe(st.session_state.masterfile)
    col1, col2 = st.columns(2, vertical_alignment="bottom")

    total_accounts = st.session_state.masterfile.height
    old_endo_count = st.session_state.masterfile.filter(pl.col("ENDO STATUS") == "OLD ENDO").height
    new_endo_count = st.session_state.masterfile.filter(pl.col("ENDO STATUS") == "NEW ENDO").height

    col1.write(f''':green[:green-background[Total Accounts: {total_accounts}]]''')
    col1.write(f''':blue[:blue-background[Old Endo: {old_endo_count}]]''')
    col1.write(f''':rainbow[:rainbow-background[New Endo: {new_endo_count}]]''')

    if st.session_state.masterfile_excel_file is None:
        st.session_state.masterfile_excel_file = save_xlsx(st.session_state.masterfile, masterfile_formatting)

    col2.download_button(
        label = f''':rainbow[Download Masterfile]''',
        data = st.session_state.masterfile_excel_file,
        file_name = f"maya_mcc_all_endo_{st.session_state.endo_date.strftime("%m%d%y")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if st.session_state.notinendo is not None:
    
    col2.download_button(
        label = "Download Not in Endo",
        data = save_xlsx(st.session_state.notinendo, masterfile_formatting),
        file_name = f"maya_mcc_not_in_endo_{st.session_state.endo_date.strftime("%m%d%y")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )