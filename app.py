import streamlit as st
from utils.process_mr import process_mr
import pandas as pd

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Clear output button
if st.button("Clear Output"):
    st.session_state.output_text = ""
    st.session_state.previous_name = ""

# Initialize session state variables if not already initialized
if "output_text" not in st.session_state:
    st.session_state.output_text = ""
if "previous_name" not in st.session_state:
    st.session_state.previous_name = ""

# Select your folder with MR
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

# Check if new files are uploaded and clear the session state accordingly
if monthly_reports:
    # Reset session state when files are uploaded
    st.session_state.output_text = ""
    st.session_state.previous_name = ""

    file_extmytime = st.file_uploader("Upload hours")
    if file_extmytime:
        results = process_mr(monthly_reports, file_extmytime)

        for index, row in results.iterrows():
            if st.session_state.previous_name != row["Name"]:
                st.session_state.previous_name = row["Name"]
                st.session_state.output_text += f"\n**Contract {row['Reference']}, {row['Name']}** \n"
                
            st.session_state.output_text += f"- {row['Error']} \n"

        st.markdown(st.session_state.output_text)
            

