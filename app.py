import streamlit as st
from utils.process_mr import process_mr
import pandas as pd


# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Select your folder with MR
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

widget = st.empty()

if monthly_reports:
    # Clear the session state to ensure fresh output each time
    st.session_state.output_text = ""
    st.session_state.previous_name = ""

    file_extmytime = st.file_uploader("Upload hours")
    if file_extmytime:
        results = process_mr(monthly_reports, file_extmytime)
        #previous_name = ""
        #text = ""
        for index, row in results.iterrows():
            if st.session_state.previous_name != row["Name"]:
                st.session_state.previous_name = row["Name"]
                st.session_state.output_text += f"\n**Contract {row['Reference']}, {row['Name']}** \n"
                
            st.session_state.output_text += f"- {row['Error']} \n"
        widget.write(str(st.session_state.output_text))    
        #st.markdown(st.session_state.output_text)

