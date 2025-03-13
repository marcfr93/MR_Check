import streamlit as st
from utils.process_mr import process_mr
import pandas as pd


def main():
    st.cache_data.clear()
    st.cache_resource.clear()
    monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True, key="reports")
    file_extmytime = st.file_uploader("Upload hours", key="hours")
    if monthly_reports and file_extmytime:
        results = process_mr(monthly_reports, file_extmytime)
        st.dataframe(results, key="results_wdg")

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")
st.cache_data.clear()
st.cache_resource.clear()
main()

        


    
