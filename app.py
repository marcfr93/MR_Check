import streamlit as st
from utils.process_mr import process_mr
import pandas as pd


def main():
    st.title("Monthly Report Review")

    st.cache_data.clear()
    st.cache_resource.clear()

    cont1 = st.container(border=True)
    with cont1:
        st.subheader("1. Upload Monthly Reports")
        monthly_reports = st.file_uploader("Upload one or more monthly reports", accept_multiple_files=True, key="reports")
    cont2 = st.container(border=True)
    with cont2:
        st.subheader("2. Upload ExtMyTime hours")
        file_extmytime = st.file_uploader("Upload single file with the ExtMyTime hours", key="hours")
    if monthly_reports and file_extmytime:
        cont3 = st.container(border=True)
        results = process_mr(monthly_reports, file_extmytime)
        with cont3:
            st.subheader("3. Review Results")
            st.dataframe(results, key="results_wdg")

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")
st.cache_data.clear()
st.cache_resource.clear()
main()

        


    
