import streamlit as st
from utils.process_mr import process_mr
import pandas as pd
#from utils.pre_process import pre_process
#from utils.process_extmytime import process_extmytime

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Select your folder with MR
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

if monthly_reports:
    file_extmytime = st.file_uploader("Upload hours")
    if file_extmytime:
        results = process_mr(monthly_reports, file_extmytime)
        for report in monthly_reports:
            st.write(f"Filename {report.name}")
        previous_name = ""
        for index, row in results.iterrows():
            if previous_name != row["Name"]:
                st.write(f"{row["Reference"], row["Name"]}")
            st.write(f"    {row["Error"]}")
            
            

