import streamlit as st
from utils.process_mr import process_mr
import pandas as pd


def write_issues(results):
    previous_name = None
    for index, row in results.iterrows():
        if previous_name != row["Name"]:
            st.write(f"\n**Contract {row['Reference']}, {row['Name']}** \n")
            previous_name = row["Name"]
        st.write(f"- {row['Error']} \n")
    return

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Clear the session state to ensure fresh output each time
if 'monthly_reports' in st.session_state:
    del st.session_state['monthly_reports']

# Select your folder with MR
st.session_state.monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

if st.session_state.monthly_reports:
    # Clear the session state to ensure fresh output each time
    if 'output_text' in st.session_state:
        del st.session_state['output_text']
    if 'previous_name' in st.session_state:
        del st.session_state['previous_name']
    if 'results' in st.session_state:
        del st.session_state['results']

    file_extmytime = st.file_uploader("Upload hours")
    if file_extmytime:
        st.session_state.results = process_mr(st.session_state.monthly_reports, file_extmytime)
        #text = ""
        #write_issues(st.session_state.results)
        st.write(st.session_state.results)
        


    
