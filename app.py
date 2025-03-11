import streamlit as st
from utils.process_mr import process_mr
import pandas as pd

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Initialize session state variables if not already initialized
if "output_text" not in st.session_state:
    st.session_state.output_text = ""
if "previous_name" not in st.session_state:
    st.session_state.previous_name = ""
if "processed" not in st.session_state:
    st.session_state.processed = False

# Clear output button
if st.button("Clear Output"):
    st.session_state.output_text = ""
    st.session_state.previous_name = ""
    st.session_state.processed = False

# Select your folder with MR
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True, key="mr_uploader")

# Check if new files are uploaded and clear the session state accordingly
if monthly_reports:
    file_extmytime = st.file_uploader("Upload hours", key="time_uploader")
    if file_extmytime:
        # Check if files have changed by comparing with stored session state
        if not st.session_state.processed or st.session_state.previous_name != monthly_reports[0].name:
            # Reset session state before processing
            st.session_state.output_text = ""
            st.session_state.previous_name = monthly_reports[0].name
            st.session_state.processed = False
            
            # Process the files
            results = process_mr(monthly_reports, file_extmytime)
            for index, row in results.iterrows():
                if st.session_state.previous_name != row["Name"]:
                    st.session_state.previous_name = row["Name"]
                    st.session_state.output_text += f"\n**Contract {row['Reference']}, {row['Name']}** \n"
                    
                st.session_state.output_text += f"- {row['Error']} \n"
            
            # Mark as processed to prevent reprocessing on rerun
            st.session_state.processed = True

    # Display the output
    st.markdown(st.session_state.output_text)

