import streamlit as st
import polars as pl
import os
import urllib.parse
from dotenv import load_dotenv
from sqlalchemy import create_engine
from datetime import datetime
from io import BytesIO

def save_excel(_df: pl.DataFrame) -> bytes:
    output = BytesIO()
    if _df is not None:
        _df.write_excel(
            output, 
            autofit=True,
            column_formats={"ACCOUNT NUM": "0"})
    output.seek(0)
    return output.getvalue()

load_dotenv()

username = os.getenv("USER_NAME")
password = os.getenv("PASSWORD")
password = urllib.parse.quote_plus(password)
hostname = os.getenv("HOSTNAME")
port = os.getenv("PORT")
database_name = os.getenv("DB_NAME")

with open("./resources/maya_active_cms.sql", "r") as sql_file:
    query = sql_file.read()

st.set_page_config(
    page_title="MAYA",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="💸"
)

if 'maya_active' not in st.session_state:
    st.session_state.maya_active = None

if 'maya_active_count' not in st.session_state:
    st.session_state.maya_active_count = None

if 'maya_active_ob' not in st.session_state:
    st.session_state.maya_active_ob = None
  
if 'maya_last_refresh' not in st.session_state:
    st.session_state.maya_last_refresh = "Last Refresh:"

col1, col2, col3, col4 = st.columns(4, vertical_alignment="bottom")

col1.title("MAYA")
refresh = col1.button("Refresh")


if refresh:
    engine = create_engine(f'mysql+mysqldb://{username}:{password}@{hostname}:{port}/{database_name}')
    maya_active = pl.read_database(query, engine)

    st.session_state.maya_active = maya_active
    st.session_state.maya_active_count = f"{maya_active.height:,}"
    st.session_state.maya_active_ob = f"{maya_active["total_debt"].sum():,.2f}"

    now = datetime.now()
    st.session_state.maya_last_refresh = f"Last Refresh: {now.strftime("%m/%d/%Y %H:%M:%S")}"
    #col1.write(st.session_state.maya_last_refresh)

col1.write(f''':green[:green-background[{st.session_state.maya_last_refresh}]]''')

cont1 = col3.container(height=150, border=True)
cont2 = col4.container(height=150, border=True)

cont1.markdown("Total Active Accounts")
cont2.markdown("Total Amount Due")
cont1.title(st.session_state.maya_active_count)
cont2.title(st.session_state.maya_active_ob)

#<<<<<<< HEAD
col_1, col_2 = st.columns([7, 1], vertical_alignment="bottom")
tagging = col_1.selectbox("Status", ["Active", "Active + POUT"])
    
if tagging == "Active + POUT":
    st.dataframe(st.session_state.maya_active_pout, height=500, use_container_width=True)
    output = st.session_state.maya_active_pout
    filename = "maya_active_pout"

if tagging == "Active":
    st.dataframe(st.session_state.maya_active, height=500, use_container_width=True)
    output = st.session_state.maya_active
    filename = "maya_active"

now = datetime.now()
col_2.download_button(
    label = "Download .xlsx",
    data = save_excel(output),
    file_name = f"{filename}_{now.strftime("%m%d%y%H%M")}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

#=======
#st.dataframe(st.session_state.maya_active, height=500, use_container_width=True)
#output = st.session_state.maya_active
#filename = "maya_active"
#>>>>>>> 8b13da970f4526cf7ffb7535199d69a5e7673872
