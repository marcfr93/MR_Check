import streamlit as st
from utils.process_mr import process_mr
import pandas as pd


# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Clear the session state to ensure fresh output each time
if 'output_text' in st.session_state:
    del st.session_state['output_text']
if 'previous_name' in st.session_state:
    del st.session_state['previous_name']
if 'monthly_reports' in st.session_state:
    del st.session_state['monthly_reports']
if 'file_extmytime' in st.session_state:
    del st.session_state['file_extmytime']

# Select your folder with MR
st.session_state.monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

if st.session_state.monthly_reports:
    # Clear the session state to ensure fresh output each time
    st.session_state.output_text = ""
    st.session_state.previous_name = ""

    st.session_state.file_extmytime = st.file_uploader("Upload hours")
    if st.session_state.file_extmytime:
        results = process_mr(st.session_state.monthly_reports, st.session_state.file_extmytime)
        #previous_name = ""
        #text = ""
        for index, row in results.iterrows():
            if st.session_state.previous_name != row["Name"]:
                st.session_state.previous_name = row["Name"]
                st.session_state.output_text += f"\n**Contract {row['Reference']}, {row['Name']}** \n"
                
            st.session_state.output_text += f"- {row['Error']} \n"
            
        st.markdown(st.session_state.output_text)

