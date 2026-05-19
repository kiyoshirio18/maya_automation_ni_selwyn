import pandas as pd
import polars as pl
import streamlit as st
from datetime import datetime
from io import BytesIO

st.header("Predictive Campaign Report to Import Remarks")

def split_dataframe(df: pl.DataFrame):
    max_rows = 10000
    # Get the number of rows in the dataframe
    n_rows = df.height
    
    # If the dataframe is larger than max_rows, split it
    if n_rows > max_rows:
        # List to store the smaller dataframes
        df_list = []
        
        # Split the dataframe in chunks of max_rows
        for start in range(0, n_rows, max_rows):
            end = min(start + max_rows, n_rows)
            df_chunk = df.slice(start, end - start)
            df_list.append(df_chunk)
        
        return df_list
    
    # If it's not larger, return a list with the original dataframe
    return [df]

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

def convert_to_datetime(date_string: str) -> datetime:
    return datetime.strptime(date_string, "%d/%m/%Y %I:%M:%S %p")

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

if "pd_report_saved" not in st.session_state:
    st.session_state.pd_report_saved = None

if "data_src_saved" not in st.session_state:
    st.session_state.data_src_saved = None

if "import_remarks" not in st.session_state:
    st.session_state.import_remarks = None

if "import_remarks_pl" not in st.session_state:
    st.session_state.import_remarks_pl = None

with st.form(key="pd_report"):
    pd_report_file = st.file_uploader("Upload Predictive Campaign Report", type="xlsx", accept_multiple_files=True)
    data_src_file = st.file_uploader("Data Source", type="xlsx")

    if st.form_submit_button(use_container_width=True):
        # Initialize an empty list to store the DataFrames
        pd_report_list = []
        
        # Loop through each uploaded file and append the data
        for uploaded_file in pd_report_file:
            # Read the Excel file into a DataFrame
            df = pd.read_excel(uploaded_file, sheet_name="Call History", dtype={"Phone Number": str})
            
            # Append the DataFrame to the list
            pd_report_list.append(df)
        
        # Concatenate all DataFrames in the list into one DataFrame
        pd_report = pd.concat(pd_report_list, ignore_index=True)
        pd_report = pd_report[pd_report["Final Dial Status"] != "Call transferred"]

        data_src = pd.read_excel(data_src_file)

        st.session_state.pd_report_saved = pd_report
        st.session_state.data_src_saved = data_src

if st.session_state.pd_report_saved is not None and st.session_state.data_src_saved is not None:

    col1, col2, col3, col4, col5 = st.columns(5, vertical_alignment="bottom")
    account_name = col1.selectbox("Account Name Field", options=st.session_state.data_src_saved.columns.to_list())
    account_num = col2.selectbox("Account Number Field", options=st.session_state.data_src_saved.columns.to_list())
    collector = col3.selectbox("Collector Field", options=st.session_state.data_src_saved.columns.to_list())
    campaign_id = col4.text_input("Campaign ID")

    if col5.button("Done", use_container_width=True):
        st.session_state.import_remarks = st.session_state.pd_report_saved

        st.session_state.import_remarks["Remark"] = st.session_state.import_remarks.apply(lambda x: remarks(campaign_id, x["Phone Number"], x["Final Dial Status"]), axis=1)
        st.session_state.import_remarks["Status"] = st.session_state.import_remarks["Final Dial Status"].apply(status)
        st.session_state.import_remarks["Time"] = st.session_state.import_remarks["Call Date"].apply(convert_to_datetime)
        st.session_state.import_remarks = pd.merge(st.session_state.import_remarks, st.session_state.data_src_saved[[account_name, account_num, collector]], left_on="Debtor", right_on=account_name, how="left")

        #Add Empty Columns
        add_empty_columns = ["PTP Date", "Reason For Default", "Field Visit Date", "Next Call Date", "PTP Amount", "Claim Paid Amount", "Relation", "Claim Paid Date"]
        for column in add_empty_columns:
            st.session_state.import_remarks[column] = None

        #Rename Columns
        st.session_state.import_remarks = st.session_state.import_remarks.rename(columns={
            account_num: "Account Number",
            "Status": "Action Status",
            "Time": "Remark Date",
            collector: "Remark By",
            "Phone Number": "Phone No"
        })

        #Select Final Columns
        st.session_state.import_remarks = st.session_state.import_remarks[['Account Number', 'Action Status', 'Remark Date', 'PTP Date', 'Reason For Default', 'Field Visit Date', 'Remark', 'Next Call Date', 'PTP Amount', 'Claim Paid Amount', 'Remark By', 'Phone No', 'Relation', 'Claim Paid Date']]
        st.session_state.import_remarks = st.session_state.import_remarks.dropna(subset=["Account Number", "Action Status"])

        st.dataframe(st.session_state.import_remarks, hide_index=True)

        st.session_state.import_remarks_pl = split_dataframe(pl.from_pandas(st.session_state.import_remarks))

if st.session_state.import_remarks_pl is not None:
    count = 0
    for df in st.session_state.import_remarks_pl:
        count = count + 1
        output_file = BytesIO()

        df.write_excel(
            output_file,
            column_formats={
                "Account Number": "0",
                "Remark Date": "mm/dd/yyyy hh:mm:ss"
            },
            autofit=True
        )
        output_file.seek(0)

        st.download_button(
            label = f"download import remarks {count}",
            data = output_file.getvalue(),
            file_name = f"import_remarks_{campaign_id}_{count}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
