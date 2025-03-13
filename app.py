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

def main():
    monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True, key="reports")
    file_extmytime = st.file_uploader("Upload hours", key="hours")
    if monthly_reports and file_extmytime:
        if st.button("Process Reports"):
            results = process_mr(monthly_reports, file_extmytime)
            #st.dataframe(results, key="results_wdg")
            downloaded = st.download_button(
                label="Download", data=results.getvalue(), file_name="results_df.xlsx")
            if downloaded:
                st.write("File downloaded!")

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")
st.cache_data.clear()
st.cache_resource.clear()
main()

"""
# Select your folder with MR
st.session_state.monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True, key="reports")
st.write(st.session_state.monthly_reports)

if st.session_state.monthly_reports:
    file_extmytime = st.file_uploader("Upload hours", key="hours")
    if file_extmytime:
        if 'results' in st.session_state:
            del st.session_state['results']
        st.session_state.results = process_mr(st.session_state.monthly_reports, file_extmytime)
        #text = ""
        #write_issues(st.session_state.results)
        st.dataframe(st.session_state.results, key="results")
"""
        


    
