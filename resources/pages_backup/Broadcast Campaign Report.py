import polars as pl
import streamlit as st
import requests
import warnings
import json
from urllib3.exceptions import InsecureRequestWarning
from resources.excel_tools import save_xlsx

# Suppress only the InsecureRequestWarning from urllib3
warnings.simplefilter("ignore", InsecureRequestWarning)

st.header("Broadcast Campaign Report")
st.write(f''':green[:green-background[Volare Credentials]]''')

if "campaign_list" not in st.session_state:
    st.session_state.campaign_list = None
if "campaign_reports" not in st.session_state:
    st.session_state.campaign_reports = None
if "header" not in st.session_state:
    st.session_state.header = {}

def get_campaign_list(keyword, header, status):

    request_url = f"https://spmpoc.volare.cc/admin/api/campaign/list?order=%20campaign_name%20asc&having=%5B%5B%5B%22campaign_id%22%2C%22contains%22%2C%22{keyword}%22%5D%2C%22or%22%2C%5B%22campaign_vdad_exten%22%2C%22contains%22%2C%22{keyword}%22%5D%2C%22or%22%2C%5B%22campaign_name%22%2C%22contains%22%2C%22{keyword}%22%5D%2C%22or%22%2C%5B%22createdName%22%2C%22contains%22%2C%22{keyword}%22%5D%2C%22or%22%2C%5B%22total%22%2C%22contains%22%2C%22{keyword}%22%5D%5D%5D&active={status}&requireTotalCount=true&hostIp="

    try:
        get_campaign_list_response = requests.get(request_url, headers=header, verify=False)
        get_campaign_list_response.raise_for_status()  # Raises an error for bad responses (4xx or 5xx)
        return get_campaign_list_response.json()
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        return None
    
def campaign_report(header, campaign_id, report_date):

    request_url = f"https://spmpoc.volare.cc/admin/api/reports/allPredictiveCampaignReport?campaignId={campaign_id}&reportDate={report_date}&campaignType=broadcast"

    try:
        campaign_report_response = requests.get(request_url, headers=header, verify=False)
        campaign_report_response.raise_for_status()  # Raises an error for bad responses (4xx or 5xx)
        st.write(campaign_report_response.json())
        return campaign_report_response.json()
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        return None
    
col1_1, col2_1, col3_1 = st.columns(3)
col1_2, col2_2, col3_2 = st.columns(3, vertical_alignment="bottom")

authorization = col1_1.text_input("Authorization", type="password", )
volare_site = col2_1.text_input("Volare-Site")
volare_id = col3_1.text_input("Volare-UserId")
keyword = col1_2.text_input("Campaign Keyword")
campaign_status = col2_2.text_input("Campaign Status")
generate_campaign_list = col3_2.button("Generate Campaign List", use_container_width=True,)

header = {}

if generate_campaign_list:
    st.session_state.header = {
        "Authorization": authorization,
        "Volare-Site": volare_site,
        "Volare-UserId": volare_id
    }

    campaign_list = get_campaign_list(keyword, st.session_state.header, campaign_status)

    st.session_state.campaign_list = {
        item["campaign_name"]: item
        for item in campaign_list["data"]
    }

if st.session_state.campaign_list is not None:
    st.text(" ")
    st.header("Campaign Details")
    campaign_name = st.selectbox("Campaign Name", options=st.session_state.campaign_list.keys())
    col1_3, col2_3, col3_3, col4_3 = st.columns([2, 2, 1, 1])
    col1_3.metric("Campaign Name", st.session_state.campaign_list[campaign_name]["campaign_name"])
    col2_3.metric("Created Date", st.session_state.campaign_list[campaign_name]["createdDate"])
    col3_3.metric("Total Accounts", st.session_state.campaign_list[campaign_name]["total"])
    col4_3.metric("Campaign ID", st.session_state.campaign_list[campaign_name]["campaign_id"])
    col1_3.metric("Audio File", st.session_state.campaign_list[campaign_name]["survey_first_audio_file"])
    col2_3.metric("Scheduled Date", st.session_state.campaign_list[campaign_name]["scheduleAt"])
    col3_3.metric("Total Dialed", st.session_state.campaign_list[campaign_name]["total_dialed"])
    col4_3.metric("Created By", st.session_state.campaign_list[campaign_name]["createdName"])

    col1_4, col2_4 = st.columns(2, vertical_alignment="bottom")
    report_date = col1_4.date_input("Report Date")
    generate_report = col2_4.button("Generate Report", use_container_width=True)

    if generate_report:
        st.session_state.campaign_reports = campaign_report(st.session_state.header, st.session_state.campaign_list[campaign_name]["campaign_id"], report_date)
    if st.session_state.campaign_reports is not None:
        call_history = pl.DataFrame(st.session_state.campaign_reports["message"]["call_history"])
        call_history = call_history.with_columns(
            pl.col("callDate").str.strptime(pl.Datetime, "%Y-%m-%d %H:%M:%S").alias("callDate"),
            pl.col("inCallTotal").cast(pl.Int64),
            pl.lit(st.session_state.campaign_list[campaign_name]["campaign_id"]).alias("Campaign ID")
        ).sort("callDate")

        col1_5, col2_5 = st.columns(2)
        col1_5.dataframe(call_history, use_container_width=True)
        col1_5.download_button(
            label = "Download Report",
            data = save_xlsx(call_history, None),
            file_name = f"{st.session_state.campaign_list[campaign_name]["campaign_name"]}_{st.session_state.campaign_list[campaign_name]["scheduleAt"]}_{st.session_state.campaign_list[campaign_name]["createdName"]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        call_history_summary = call_history.group_by("dialStatus").agg(
            pl.col("dialStatus").len().alias("Count"),
            pl.col("inCallTotal").sum().alias("Total Duration (s)")
        ).sort("dialStatus")

        call_history_total = pl.DataFrame({
            "dialStatus": "Total",
            "Count": call_history_summary["Count"].sum(),
            "Total Duration (s)": call_history_summary["Total Duration (s)"].sum()
        })

        call_history_total = call_history_total.with_columns(
            pl.col("Count").cast(pl.UInt32),
            pl.col("Total Duration (s)").cast(pl.Int64)
        )

        call_history_summary = call_history_summary.extend(call_history_total)
        col2_5.dataframe(call_history_summary, use_container_width=True)