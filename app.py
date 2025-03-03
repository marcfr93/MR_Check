import streamlit as st
#from utils.process_mr import process_mr
#from utils.pre_process import pre_process
#from utils.process_extmytime import process_extmytime

# Configure layout of page, must be first streamlit call in script
st.set_page_config(layout="wide")

# Select your folder with MR
text = st.text_area("Select the folder with the Monthly Reports to be processed")
monthly_reports = st.file_uploader("Upload Monthly Reports", accept_multiple_files=True)


# Select the file of extmytime hours
if monthly_reports:
    text = st.text_area("Select the XLSX file with the hours from ExtMyTime")
    hours_extmytime = st.file_uploader("Upload hours")
    if hours_extmytime:
        check_mr()

"""
# add a button for the extmytime process
text = st.text_area("Copy your Extmytime here")
if text:
    total_hours, tasks_hours, message = process_extmytime(text)
    if message == "":
        st.write("Total hours:", total_hours)
        st.write("Tasks hours:", tasks_hours)
        for task_code in tasks_hours.keys():
            st.write(f"Task {task} hours: {tasks_hours[task]}")
    else:
        st.write(message)

if text and message == "":
    uploaded_file = st.file_uploader("Upload last month MR")

    downloaded = False

    if uploaded_file:
        new_mr, new_name = pre_process(uploaded_file, tasks_hours, total_hours)
        st.write("Thanks for uploading, now you can download the new MR")
        downloaded = st.download_button(
            label="Download", data=new_mr.getvalue(), file_name=new_name
        )

    if downloaded:
        st.write("File Downloaded!")
"""