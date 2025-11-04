import streamlit as st
from utils.process_mr import process_mr
from utils.process_mr_1639 import process_mr_1639
import pandas as pd


def main():
    st.title("Monthly Report Review")

    st.cache_data.clear()
    st.cache_resource.clear()

    tab1, tab2 = st.tabs(["Review 1159 Reports", "Review 1639 Reports"])

    with tab1:
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

    with tab2:
        cont1 = st.container(border=True)
        with cont1:
            st.subheader("1. Upload Monthly Reports")
            monthly_reports = st.file_uploader("Upload one or more monthly reports", accept_multiple_files=True, key="reports_1639")

        if monthly_reports: #and file_extmytime:
            cont2 = st.container(border=True)
            results = process_mr_1639(monthly_reports) #, file_extmytime)
            with cont2:
                st.subheader("2. Review Results")
                st.dataframe(results, key="results_wdg_1639")

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")
st.cache_data.clear()
st.cache_resource.clear()
main()

        


    
