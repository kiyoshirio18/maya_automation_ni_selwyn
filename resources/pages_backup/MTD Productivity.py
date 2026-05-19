import polars as pl
import streamlit as st

st.header("MTD Productivity")
st.write(f''':green[:green-background[Remark Report]] :green[:green-background[PTP Booked]]''')

if "attempts_summ" not in st.session_state:
    st.session_state.attempts_summ = None

if "ptp_summ" not in st.session_state:
    st.session_state.ptp_summ = None

col1, col2 = st.columns(2)

with col1.form(key="mtd"):
    daily_remark_file = st.file_uploader("Upload Remark File", type="xlsx")

    if st.form_submit_button("SUBMIT", use_container_width=True):
        daily_remark = pl.read_excel(daily_remark_file)

        acr_count = daily_remark.group_by("Collector").agg(
            pl.col("Account No.").n_unique().alias("ACR COUNT")
        )

        attempts = daily_remark.group_by("Collector").len(name="ATTEMPTS")

        connected_calls = daily_remark.filter(pl.col("Call Status") == "CONNECTED")

        connected_count = connected_calls.group_by("Collector").len(name="CONTACTS")

        merged = acr_count.join(attempts, on="Collector", how="inner")

        merged = merged.join(connected_count, on="Collector", how="inner")

        st.session_state.attempts_summ = merged.sort("Collector")

with col2.form(key="ptp"):
    ptp_booked_file = st.file_uploader("Upload PTP Booked File", type="xlsx")

    if st.form_submit_button("SUBMIT", use_container_width=True):
        ptp_booked = pl.read_excel(ptp_booked_file)

        ptp_total = ptp_booked.group_by("Mediator").agg(
            pl.col("Account_Number").len().alias("Total PTP"),
            pl.col("PTP Amount").sum().alias("Total PTP Amount")
        ).sort("Mediator")

        due_ptp = ptp_booked.filter(pl.col("STATUS") == "PTP").group_by("Mediator").agg(
            pl.col("Account_Number").len().alias("Due PTP"),
            pl.col("PTP Amount").sum().alias("Due PTP Amount")
        ).sort("Mediator")

        ptp_kept = ptp_booked.filter((pl.col("STATUS") == "PAID") | (pl.col("STATUS") == "PAID (PARTIAL)")).group_by("Mediator").agg(
            pl.col("Account_Number").len().alias("PTP Kept"),
            pl.col("PTP Amount").sum().alias("Kept Amount")
        ).sort("Mediator")

        broken_ptp = ptp_booked.filter(pl.col("STATUS") == "BP").group_by("Mediator").agg(
            pl.col("Account_Number").len().alias("Broken PTP"),
            pl.col("PTP Amount").sum().alias("Broken PTP Amount")
        ).sort("Mediator")

        ptp_summary = ptp_total.join(due_ptp, on="Mediator", how="left")
        ptp_summary = ptp_summary.join(ptp_kept, on="Mediator", how="left")
        ptp_summary = ptp_summary.join(broken_ptp, on="Mediator", how="left")
        ptp_summary = ptp_summary.fill_null(strategy="zero")

        st.session_state.ptp_summ = ptp_summary

if st.session_state.attempts_summ is not None:
    col1.dataframe(st.session_state.attempts_summ, use_container_width=True)

if st.session_state.ptp_summ is not None:
    col2.dataframe(st.session_state.ptp_summ, use_container_width=True)