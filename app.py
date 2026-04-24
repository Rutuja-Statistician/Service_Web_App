import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Import your functions from main.py
from main import (
    func1, circlewise_platter, statuswise_platter, 
    billing_code_status_platter, pdna_status_platter, 
    ran_cn_due_status_platter, dealerwise_status_platter
)

st.set_page_config(page_title="Service Report Generator", layout="wide")

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to", ["Upload File & Create Report", "View Dashboard"])

if page == "Upload File & Create Report":
    st.header("Upload Raw Data")
    
    # 1. File Uploader
    uploaded_raw_file = st.file_uploader("Choose the Raw Data Excel file", type=["xlsx"])


    if uploaded_raw_file is not None:
        if st.button("Generate Report"):
            with st.spinner("Processing data and applying formatting..."):
                try:
                    # 2. Run your processing logic
                    final_df = func1(uploaded_raw_file)

                    if isinstance(final_df, pd.DataFrame):
                        # 3. Create an In-Memory Buffer for the Excel file
                        # This allows users to download the file from the web interface
                        output = io.BytesIO()
                        
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            circlewise_platter(final_df, writer)
                            statuswise_platter(final_df, writer)
                            billing_code_status_platter(final_df, writer)
                            pdna_status_platter(final_df, writer)
                            ran_cn_due_status_platter(final_df, writer)
                            dealerwise_status_platter(final_df, writer)
                        
                        processed_data = output.getvalue()

                        # 4. Success UI
                        st.success("Report Generated Successfully!")
                        
                        # Show a preview
                        st.subheader("Data Preview (Processed)")
                        
                        # 5. Download Button
                        st.download_button(
                            label="📥 Download Platter Report",
                            data=processed_data,
                            file_name=f"service_platter_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.dataframe(
                            final_df,
                            use_container_width=True,
                            column_config={"Select": st.column_config.CheckboxColumn()} # if you need selection
                        )
                    
                except Exception as e:
                    st.error(f"An error occurred: {e}")
