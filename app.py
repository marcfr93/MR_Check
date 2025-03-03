import streamlit as st
#from utils.process_mr import process_mr
#from utils.pre_process import pre_process
#from utils.process_extmytime import process_extmytime

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Select your folder with MR
text = st.text_area("Select the folder with the Monthly Reports to be processed")
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

if monthly_reports:
    text = st.text_area("Select the XLSX file with the hours from ExtMyTime")
    hours_extmytime = st.file_uploader("Upload hours")
    if hours_extmytime:
        check_mr()

