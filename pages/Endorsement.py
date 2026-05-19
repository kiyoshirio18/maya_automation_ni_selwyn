import polars as pl
import streamlit as st
import json
import msoffcrypto
import pickle
import os
import re
from io import BytesIO
from datetime import datetime
from resources.excel_tools import cast_columns, save_xlsx, xlsx_to_xls

st.header("Endorsement")
st.write(f''':green[:green-background[Endorsement Raw Files]] :green[:green-background[Masterfile]]''')

def clean_column_names(df: pl.DataFrame) -> pl.DataFrame:
    """
    Removes trailing underscore + number (e.g. '_1', '_23') from column names in a Polars DataFrame.

    Example:
        'mobile_number_1' -> 'mobile_number'
    """
    new_columns = {}
    for col in df.columns:
        new_name = re.sub(r'_\d+$', '', col)
        new_columns[col] = new_name
    
    return df.rename(new_columns)

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
            tag_list = ["PJJL", "PJJB", "PJCM", "PKGV", "PJGF", "PAHD", "PAOB", "PJDX", "PGDJ", "PRIP", "PMQR", "PJVY", "PDUM", "PRJL", "PNPL"]
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

with open("./resources/maya_negosyo_mapping.json", "r") as file:
    maya_negosyo_mapping = json.load(file)

with open("./resources/maya_credit_mapping.json", "r") as file:
    maya_credit_mapping = json.load(file)

with open("./resources/maya_sme_mapping.json", "r") as file:
    maya_sme_mapping = json.load(file)

with open("./resources/maya_product_placement.json", "r") as file:
    maya_product_placement = json.load(file)

with open("./resources/masterfile_datatypes.pkl", "rb") as file:
    masterfile_datatypes = pickle.load(file)

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

with st.form(key="endorsement"):
    col1, col2 = st.columns(2)

    credit = col1.file_uploader("Maya Credit Endo Raw File", type="xlsx")
    negosyo = col1.file_uploader("Maya Negosyo Endo Raw File", type="xlsx")
    sme = col2.file_uploader("Maya SME Endo Raw File", type="xlsx")

    masterfile = col2.file_uploader("Yesterday's Masterfile", type="xlsx")
    endo_date = col1.date_input("Endorsement Date", value="today")
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
                        pl.lit("OLD ENDO").alias("ENDO STAT")
                    )
                    
                    st.session_state.masterfile_yesterday = masterfile_yesterday
                except msoffcrypto.exceptions.DecryptionError as e:
                    if str(e) == "Unencrypted document" or str(e) == "No key specified":
                        print("Caught DecryptionError with message: Unencrypted document")
                        masterfile_yesterday = pl.read_excel(masterfile)
                        masterfile_yesterday = cast_columns(masterfile_yesterday, masterfile_datatypes)
                        st.session_state.masterfile_yesterday = masterfile_yesterday
                    else:
                        print(f"Caught DecryptionError with a different message: {e}")

                st.write("Aligning Raw File Columns to Mastefile Headers...")
                files = {
                    'negosyo_raw_file': negosyo,
                    'credit_raw_file': credit,
                    'sme_raw_file': sme
                }

                for var_name, file in files.items():
                    if file is None:
                        globals()[var_name] = None
                    else:
                        globals()[var_name] = pl.read_excel(file, schema_overrides={
                            "ACCOUNT_ID": pl.Utf8,
                            "ACCOUNT_NUMBER_LAST_SET_TO_ARREARS_DATE": pl.Utf8
                        })
                
                # Dictionary to map raw files to their respective mappings
                raw_files = {
                    "negosyo": {"file": globals().get("negosyo_raw_file"), "mapping": maya_negosyo_mapping},
                    "credit": {"file": globals().get("credit_raw_file"), "mapping": maya_credit_mapping},
                    "sme": {"file": globals().get("sme_raw_file"), "mapping": maya_sme_mapping}  # Assuming sme_raw_file and maya_sme_mapping exist
                }
#<<<<<<< HEAD

#=======
                
#>>>>>>> 8b13da970f4526cf7ffb7535199d69a5e7673872
                if raw_files["negosyo"]["file"] is not None:
                    raw_files["negosyo"]["file"] = clean_column_names(raw_files["negosyo"]["file"])

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
                        if file_name == "negosyo":
                            negosyo_masterfile = masterfile
                        elif file_name == "credit":
                            credit_masterfile = masterfile
                        elif file_name == "sme":
                            sme_masterfile = masterfile
                    else:
                        # Set the corresponding masterfile to None if raw_file is None
                        if file_name == "negosyo":
                            negosyo_masterfile = None
                        elif file_name == "credit":
                            credit_masterfile = None
                        elif file_name == "sme":
                            sme_masterfile = None
            
            st.write("Stacking Aligned Endo Files...")
            # List of masterfile variables
            masterfiles = [credit_masterfile, negosyo_masterfile, sme_masterfile]

            # Initialize the final masterfile as None
            masterfile = None

            # Loop through the masterfiles and stack them if they exist
            for file in masterfiles:
                if file is not None:  # Check if the masterfile exists
                    if masterfile is None:
                        masterfile = file  # Initialize masterfile with the first valid one
                    else:
                        masterfile = masterfile.vstack(file)

            st.write("Filling in Necessary Columns...")
            #Filling in necessary columns
            masterfile = masterfile.with_columns(
                (pl.col("FIRST_NAME") + " " + pl.col("LAST_NAME")).str.to_uppercase().alias("NAME"),
                pl.col("DPD_").map_elements(get_dpd_bucket, return_dtype=pl.Utf8).alias("DPD BUCKET"),
                (pl.col("OB") + (pl.col("OB") * 0.0017)).alias("OB adjustments"),
                pl.col("MOBILE_NUMBER_1").map_elements(mobile_number, return_dtype=pl.Utf8).alias("MOBILE PROPER"),
                pl.when(pl.col("ACCOUNT NUMBER").is_in(masterfile_yesterday["ACCOUNT NUMBER"])).then(pl.lit("OLD ENDO")).otherwise(pl.lit("NEW ENDO")).alias("ENDO STAT")
            )

            masterfile = masterfile.with_columns(
                pl.col(f"TU_NUMBER_{x}").map_elements(mobile_number, return_dtype=pl.Utf8).alias(f"TU_NUMBER_{x}") for x in range(1,6)
            )

            masterfile = masterfile.with_columns(
                (pl.col("PRODUCT_NAME").replace_strict(maya_product_placement, default=None) + " " + pl.col("DPD BUCKET")).alias("PLACEMENT")
            )

            # List of columns to be updated
            columns_to_update = ["CHCODE", "TAGGING", "RECEIVED DATE"]

            # Loop through the columns to update
            for col in columns_to_update:
                # Perform the join for each column
                masterfile = masterfile.join(
                    masterfile_yesterday.select(["ACCOUNT NUMBER", col]), 
                    on="ACCOUNT NUMBER", 
                    how="left", 
                    suffix="_new"  # No need for .lower()
                )

                # Update the column with the new values if they exist in the joined DataFrame
                masterfile = masterfile.with_columns(
                    pl.when(pl.col(f"{col}_new").is_not_null())
                    .then(pl.col(f"{col}_new"))
                    .otherwise(pl.col(col))
                    .alias(col)
                )

            #Assign taggings for new endo accounts
            tag_assigner = TagAssigner()
            masterfile = masterfile.with_columns(
                pl.when(
                    pl.col("ENDO STAT") == "NEW ENDO"
                    ).then(
                        pl.col("PLACEMENT").map_elements(tag_assigner.assign_taggings, return_dtype=pl.Utf8).alias("TAGGING")
                    ).otherwise(
                        pl.col("TAGGING")
                    )
            )
            tag_assigner.save_index()

            masterfile = masterfile.with_columns(
                pl.when(
                    pl.col("ENDO STAT") == "NEW ENDO"
                    ).then(
                        pl.lit(endo_date).alias("RECEIVED DATE")
                    ).otherwise(
                        pl.col("RECEIVED DATE")
                    )
            )

            masterfile = masterfile.with_columns(
                pl.col("RECEIVED DATE").map_elements(check_date_status, return_dtype=pl.Utf8).alias("FRESH/SPILLOVER")
            )

            masterfile = masterfile.drop(["DPD", "CHCODE_new", "TAGGING_new", "RECEIVED DATE_new"])

            # List of masterfile variables
            masterfiles = [credit_masterfile, negosyo_masterfile, sme_masterfile]
            
            maya_products = []

            for product in maya_product_placement.keys():
                maya_products.append(product)

            new_endo_products = masterfile["PRODUCT_NAME"].unique().to_list()

            if set(maya_products) != set(new_endo_products):
                non_present_product = [product for product in maya_products if product not in new_endo_products]
                non_present_product_df = masterfile_yesterday.filter(pl.col("PRODUCT_NAME").is_in(non_present_product))
                masterfile = masterfile.vstack(non_present_product_df)

            st.session_state.masterfile = masterfile.sort("PLACEMENT").unique(keep="any")
            st.session_state.notinendo = st.session_state.masterfile_yesterday.join(st.session_state.masterfile, on="ACCOUNT NUMBER", how="anti")
            status.update(label="Done", state="complete", expanded=False)


if st.session_state.masterfile is not None:
    st.dataframe(st.session_state.masterfile)
    col1, col2 = st.columns(2, vertical_alignment="bottom")

    total_accounts = st.session_state.masterfile.height
    old_endo_count = st.session_state.masterfile.filter(pl.col("ENDO STAT") == "OLD ENDO").height
    new_endo_count = st.session_state.masterfile.filter(pl.col("ENDO STAT") == "NEW ENDO").height

    col1.write(f''':green[:green-background[Total Accounts: {total_accounts}]]''')
    col1.write(f''':blue[:blue-background[Old Endo: {old_endo_count}]]''')
    col1.write(f''':rainbow[:rainbow-background[New Endo: {new_endo_count}]]''')

    if st.session_state.masterfile_excel_file is None:
        st.session_state.masterfile_excel_file = save_xlsx(st.session_state.masterfile, masterfile_formatting)

    col2.download_button(
        label = f''':rainbow[Download Masterfile]''',
        data = st.session_state.masterfile_excel_file,
        file_name = f"maya_all_endo_{datetime.now().strftime("%m%d%y")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if st.session_state.notinendo is not None:
    
    col2.download_button(
        label = "Download Not in Endo",
        data = save_xlsx(st.session_state.notinendo, masterfile_formatting),
        file_name = f"maya_not_in_endo_{datetime.now().strftime("%m%d%y")}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )