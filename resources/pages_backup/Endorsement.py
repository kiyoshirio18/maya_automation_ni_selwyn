import polars as pl
import streamlit as st
import json
import msoffcrypto
import pickle
from datetime import datetime
from resources.excel_tools import cast_columns

st.header("Endorsement")

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

with open("maya_negosyo_mapping.json", "r") as file:
    maya_negosyo_mapping = json.load(file)

with open("maya_credit_mapping.json", "r") as file:
    maya_credit_mapping = json.load(file)

with open("maya_sme_mapping.json", "r") as file:
    maya_sme_mapping = json.load(file)

with open("maya_product_placement.json", "r") as file:
    maya_product_placement = json.load(file)

with open("masterfile_datatypes.pkl", "rb") as file:
    masterfile_datatypes = pickle.load(file)

with open("maya_masterfile_formatting.json", "r") as file:
    masterfile_formatting = json.load(file)