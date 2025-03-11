import streamlit as st
from utils.process_mr import process_mr
import pandas as pd

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Check if session state variables are initialized
if "output_text" not in st.session_state:
    st.session_state.output_text = ""
if "previous_name" not in st.session_state:
    st.session_state.previous_name = ""

# Select your folder with MR
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

# Clear the output if new files are uploaded
if monthly_reports:
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
            

