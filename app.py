import streamlit as st
import pandas as pd
import io
from datetime import datetime

from main import (
    func1,
    circlewise_platter,
    statuswise_platter,
    billing_code_status_platter,
    pdna_status_platter,
    ran_cn_due_status_platter,
    dealerwise_platter,          # ← make sure this matches main.py exactly
    fetch_and_format_report
)

from src.sidebar import render_sidebar

st.set_page_config(page_title="Service Report Generator", layout="wide")

page = render_sidebar()

# --- PAGE: Upload File & Create Report ---
if page == "upload":
    st.header("📤 Upload File & Create Report")

    uploaded_raw_file = st.file_uploader("Choose the Raw Data Excel file", type=["xlsx"])

    if uploaded_raw_file is not None:
        if st.button("Generate Report"):
            with st.spinner("Processing data and pushing to Database..."):
                try:
                    # Step 1: Process raw file → push all sheets to Google Sheets
                    final_df = func1(uploaded_raw_file)

                    if isinstance(final_df, pd.DataFrame):
                        circlewise_platter(final_df)
                        statuswise_platter(final_df)
                        billing_code_status_platter(final_df)
                        pdna_status_platter(final_df)
                        ran_cn_due_status_platter(final_df)
                        dealerwise_platter(final_df)

                        st.success("✅ All data updated in Database!")

                except Exception as e:
                    st.error(f"Error during processing: {e}")

    st.divider()

    # --- Download Section (independent of upload) ---
    st.subheader("📥 Download Platter Report")
    st.caption("Fetches latest data directly from Database and downloads as Excel.")

    if st.button("📥 Download Platter Report"):
        with st.spinner("Fetching data from Database and formatting..."):
            report_data = fetch_and_format_report()

            if report_data:
                st.download_button(
                    label="⬇️ Click here to Download",
                    data=report_data,
                    file_name=f"service_platter_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    st.divider()

    # --- Download Section (independent of upload) ---
    st.subheader("📥 Email Reports")

# --- PAGE: View Dashboard ---
elif page == "dashboard":
    st.header("📊 View Dashboard")
    st.info("Dashboard coming soon...")

# --- PAGE: View Report ---
elif page == "reports":
    st.header("📌 View Report")
    st.info("Reports coming soon...")
