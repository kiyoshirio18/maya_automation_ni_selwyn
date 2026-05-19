import polars as pl
import streamlit as st
from datetime import datetime
from io import BytesIO

def remarks(campaign_id, number, status):


    if status == "No Answer(NA)":
        return f"Predictive {campaign_id} called {number}: predictive call automatically detected as no answer"
    
    if status == "Auto Busy":
        return f"Predictive {campaign_id} called {number}: predictive call automatically detected as busy"
    
    if status == "ADC/DC":
        return f"Predictive {campaign_id} called {number}: predictive call automatically detected as disconnected"
    
    if status == "Dropped":
        return f"Predictive {campaign_id} called {number}: predictive call connected, but debtor hung up before being transferred to collector"
    
    return None

def status(status):

    if status == "No Answer(NA)":
        return f"RNA"
    
    if status == "Auto Busy":
        return f"BUSY"
    
    if status == "ADC/DC":
        return f"UNTC"
    
    if status == "Dropped":
        return f"DROPPED"
    
    return None

st.header("Predictive Campaign Report to Daily Remark")

if "daily_remark_file" not in st.session_state:
    st.session_state.daily_remark_file = None
if "file_date" not in st.session_state:
    st.session_state.file_date = None

with st.form(key="pd_report_daily_remark"):
    pd_report_file = st.file_uploader("Predictive Campaign Report", accept_multiple_files=True, type="xlsx")
    data_grid_file = st.file_uploader("Data Grid", type="xlsx")
    
    col1, col2 = st.columns(2, vertical_alignment="bottom")
    campaign_id = col1.text_input("Campaign ID", placeholder="CP"),

    if col2.form_submit_button(use_container_width=True):
        # Initialize an empty list to store the DataFrames
        pd_report_list = []
        
        # Loop through each uploaded file and append the data
        for uploaded_file in pd_report_file:
            # Read the Excel file into a DataFrame
            df = pl.read_excel(uploaded_file, engine="openpyxl", sheet_name="Call History", schema_overrides={"Phone Number": pl.Utf8})
            
            # Convert all columns to string type (Utf8)
            df = df.with_columns([pl.col(col).cast(pl.Utf8) for col in df.columns])
            
            # Append the DataFrame to the list
            pd_report_list.append(df)
        
        # Concatenate all DataFrames in the list into one DataFrame
        pd_report = pl.concat(pd_report_list, how="vertical")

        #Load Data Grid file to Data Frame
        data_grid = pl.read_excel(data_grid_file, engine="openpyxl")

        #Exclude Connected Calls
        pd_report = pd_report.filter(pl.col("Final Dial Status") != "Call transferred")

        #Apply Functions
        campaign_id = "CP16944"
        pd_report = pd_report.with_columns(
            pl.col("Call Date").str.to_datetime(format="%d/%m/%Y %I:%M:%S %p").alias("Time"),
            pl.col("Final Dial Status").map_elements(status, return_dtype=pl.Utf8).alias("Status"),
            pl.struct(["Phone Number", "Final Dial Status"]).map_elements(lambda x: remarks(campaign_id, x["Phone Number"], x["Final Dial Status"]), return_dtype=pl.Utf8).alias("Remark")
        )

        #Join Necessary Fields from Data Grid
        pd_report = pd_report.join(data_grid["Name", "Account No.", "Days Past Due (DPD)", "Collector", "Client Name", "Product Type", "Batch No.", "Balance"], left_on="Debtor", right_on="Name", how="left")

        #Add Necessary Columns
        pd_report = pd_report.with_columns(
            pl.col("Time").dt.date().alias("Date"),
            pl.lit("Predictive").alias("Remark Type"),
            pl.col("Collector").alias("Remark By"),
            pl.col("Balance").cast(pl.Float64)
        )

        add_columns = ['S.No', 'Card No.', 'Service No.', 'Reason For Default', 'Call Status', 'Field Visit Date', 'Product Description', 'Account Type', 'Relation', 'PTP Amount', 'Next Call', 'PTP Date', 'Claim Paid Amount', 'Claim Paid Date', 'Days Past Write Off', 'Contact Type', 'Black Case No.', 'Red Case No.', 'Court Name', 'Lawyer', 'Legal Stage', 'Legal Status', 'Next Legal Follow up', 'Call Duration', 'Talk Time Duration']

        pd_report = pd_report.with_columns(
            pl.lit(None).alias(column) for column in add_columns
        )
        
        #Rename some columns based on Daily Remark Header
        pd_report = pd_report.rename({
            "Days Past Due (DPD)": "DPD",
            "Client Name": "Client",
            "Batch No.": "Batch No",
            "Phone Number": "Dialed Number"
        })

        #Select Necessary Columns and filter out Null Account No. Values
        daily_remark = pd_report.select([
            "S.No", "Date", "Time", "Debtor", "Account No.", "Card No.", "Service No.",
            "DPD", "Reason For Default", "Call Status", "Status", "Remark", "Remark By", 
            "Remark Type", "Field Visit Date", "Collector", "Client", "Product Description", 
            "Product Type", "Batch No", "Account Type", "Relation", "PTP Amount", "Next Call", 
            "PTP Date", "Claim Paid Amount", "Claim Paid Date", "Dialed Number", "Days Past Write Off", 
            "Balance", "Contact Type", "Black Case No.", "Red Case No.", "Court Name", "Lawyer", 
            "Legal Stage", "Legal Status", "Next Legal Follow up", "Call Duration", "Talk Time Duration"
        ]).filter(pl.col("Account No.").is_not_null())

        #Write Excel File to Bytes
        daily_remark_bytes = BytesIO()

        daily_remark.write_excel(
            daily_remark_bytes,
            dtype_formats = {
                pl.Date: "mm/dd/yyyy",
                pl.Datetime: "mm/dd/yyyy hh:mm:ss",
                pl.Float64: "0.00"
            },
            autofit=True
        )

        daily_remark_bytes.seek(0)
        st.session_state.daily_remark_file = daily_remark_bytes
        st.session_state.file_date = daily_remark["Date"].max()

if st.session_state.daily_remark_file is not None:
    st.download_button(
            label = f"download daily remarks",
            data = st.session_state.daily_remark_file.getvalue(),
            file_name = f"daily_remarks_{st.session_state.file_date.strftime("%m%d%y")}_additional.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )