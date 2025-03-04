import streamlit as st
from utils.process_mr import process_mr
import pandas as pd

# st.write(f"**Contract {row["Reference"]}, {row["Name"]}**")
# st.markdown(f"- {row["Error"]} ")

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Select your folder with MR
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)

if monthly_reports:
    file_extmytime = st.file_uploader("Upload hours")
    if file_extmytime:
        results = process_mr(monthly_reports, file_extmytime)
        previous_name = ""
        text = ""
        for index, row in results.iterrows():
            if previous_name != row["Name"]:
                previous_name = row["Name"]
                text = text + f"\n**Contract {row["Reference"]}, {row["Name"]}** \n"
                
            text = text + f"- {row["Error"]} \n"
            
        st.markdown(text)
            
            

