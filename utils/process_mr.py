import os
import re
from datetime import datetime
from pathlib import Path
import warnings
import docx2txt
import docx
import pandas as pd
import math
from unidecode import unidecode
import time


FOLDER = "test_data"
MONTH_NUMBER_TO_NAME = {
    1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December",
}

def read_info_files(folder):

    # Get file with list of names
    list_names_file = Path("F4E-OMF-1159-01-01 List of Names - FINAL VERSION.xlsx")
    try:
        list_names = pd.read_excel(list_names_file, header=0)
    except FileNotFoundError:
        print(f" Could not find the list of names file {list_names_file}")

    # Get files with F4E customer references
    f4e_customer_ref_path = Path("F4E Customer Ref.xlsx")
    try:
        f4e_customer_ref = pd.read_excel(f4e_customer_ref_path, header=0)
    except FileNotFoundError:
        print(f" Could not find the file with the F4E Customer References")

    return list_names, f4e_customer_ref

results_df = pd.DataFrame(columns=["Reference", "Name", "Error"])
list_names, f4e_customer_ref = read_info_files(FOLDER)
global hours_task_plan

def process_mr(mr_files, hours_task_plan):
    
    # results_df = pd.DataFrame(columns=["Reference", "Name", "Error"])
    # total_hours_df = pd.DataFrame(columns=["Reference", "Name", "Hours in report header"])
    
    hours_task_plan = pd.read_excel(hours_task_plan, skiprows=3)
    for report in mr_files:
        if report.name.endswith(".docx"):
            process_monthly(report, hours_task_plan)


def diff_month(d1, d2):
    """
    Returns the difference of months between two dates

    Arguments:
        d1 (datetime)
        d2 (datetime)
    Returns:
        int
    """
    return (d1.year - d2.year) * 12 + d1.month - d2.month


class HeaderData:
    def __init__(self):
        self.report_number = None
        self.f4e_reference = None
        self.supplier_dms = None
        self.kom_date = None
        self.reported_hours = None
        self.version = None
        self.customer_ref = None

    @property
    def totally_filled(self):
        return (
            self.report_number
            and self.f4e_reference
            and self.supplier_dms
            and self.kom_date
            and self.reported_hours
            and self.version
            and self.customer_ref
        )

    def __str__(self):
        text = f"""
Report number: {self.report_number}
F4E reference: {self.f4e_reference}
Supplier DMS: {self.supplier_dms}
KOM Date: {self.kom_date}
Reported hours: {self.reported_hours}
Revision: {self.version}
Customer Ref.: {self.customer_ref}"""
        return text


class Name:
    """
    A class to represent the different configuration of a person's name

    Attributes
    ---
    report: str
    surname: str
    age: str
    """
    def __init__(self):
        self.report = None
        self.irs = None
        self.irs_comma = None

    def convert(self, list_names):
        self.irs_comma = unidecode(list_names[list_names["Monthly/Mission Request"].apply(unidecode) == unidecode(self.report)]["IRS"].values[0])
        name_irs = self.irs_comma.split(", ")
        self.irs = unidecode(name_irs[1] + " " + name_irs[0])


def read_header(document):
    """
    Extracts the data from the header of the monthly report

    Parameters:
        document (str): string with the text of the report

    Returns
        HeaderData class
    """
    header_data = HeaderData()
    end = re.search("Tasks and milestones", document).span()[0]
    lines = list(filter(None, document[:end].splitlines()))
    for i, line in enumerate(lines):
        if "Report Number:" in line:
            header_data.report_number = lines[i + 1].strip()
        elif "Revision:" in line:
            header_data.version = lines[i + 1].strip()
        elif "F4E reference:" in line:
            header_data.f4e_reference = lines[i + 1].strip()
        elif "F4E Customer ref:" in line:
            header_data.customer_ref = lines[i + 1].strip()
        elif "Supplier DMS#:" in line:
            header_data.supplier_dms = lines[i + 1].strip()
        elif "KOM Date" in line:
            header_data.kom_date = lines[i + 1].strip()
        elif "Hours implemented in the period" in line:
            to_be_replaced = [" ", "(", "*", ")"]
            hours = lines[i + 1].strip().replace(",", ".")
            for j in to_be_replaced:
                hours = hours.replace(j, "")
            header_data.reported_hours = hours
        if header_data.totally_filled:
            break
    if not header_data.totally_filled:
        error_message = f"  The monthly header was not successfully read {header_data}"
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
        raise ValueError(error_message)
    return header_data


def add_header_hours_to_list(header_data):
    """
    Adds in the total_hours_df the hours done by the person during the period, as per the information
    in the report header

    Parameters:
        header_data (class)
    """
    total_hours_df.loc[len(total_hours_df)] = [header_data.f4e_reference, name_report, header_data.reported_hours]
    return


def show_version_message(header_data):
    """
    Prints a messages with the Revision number of the document
    :param header_data(class)
    :return: None
    """
    print(f"  The revision number in the header is {header_data.version}. Make sure this is correct.")
    return


def get_names(list_names, filename):
    """
    Initializes class name, reads the name in the title of the report and converts it to the other two formats.

    Parameters:
        list_names (dataframe): extracted from the file with all the names and their equivalents
        filename (str): file name of the monthly report file
    Returns:
        name (class)
    """
    name = Name()
    name.report = unidecode(re.match(r".+ Monthly Report (.+\s.+) #", filename).group(1))
    name.convert(list_names)

    return name


def check_f4e_contract(code_from_filename, header_data):
    """Check if F4E reference is the same in the name of the report and inside the report"""
    if header_data.f4e_reference not in code_from_filename:
        error_message = f"  The F4E contract shown in the header ({header_data.f4e_reference}) does not match " \
                        f"the one of the Word filename ({code_from_filename})"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    return


def check_supplier_dms(header_data, name, list_employees):
    """Check if the report of the month, year and person is in the Excel file 'DMS Number Monthly Report.xlsx' and
    also if the DMS number is correspondent"""
    dms_data_file = Path("DMS Number Monthly Report.xlsx")
    try:
        dms_table = pd.read_excel(dms_data_file, header=0)
    except FileNotFoundError:
        print(f"  Could not find the DMS data file {dms_data_file}.")
        return

    # Get month from report number in header
    month = re.match(r"#\d+_M(\d+)_\d+", header_data.report_number).group(1)  # e.g. "3"
    month = int(month)
    month = MONTH_NUMBER_TO_NAME[month]  # e.g. "March"
    # Get year from report number in header
    year = re.match(r"#\d+_M\d+_(\d+)", header_data.report_number).group(1)  # eg."2022"
    # Get name from filename

    if name.irs is None:
        error_message = f"  The name '{name.report}' could not be found in the file with the list of names under the " \
                        f"column named 'Monthly/Mission request'. The DMS number could not be checked."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
        return
    # Transform the name using the file with the list of names and their equivalents
    string_dms = f'^(?=.*{year})(?=.*{month})(?=.*{unidecode(name.irs)})|^(?=.*{year})(?=.*{month})(?=.*{unidecode(name.report)})'
    dms_table = dms_table[dms_table["Description"].apply(unidecode).str.contains(string_dms)]
    """
    person = list_employees[list_employees["Name Monthly/Mission"].apply(unidecode) == name.report]
    dms = person[f"{month} {year}"]
    if not pd.isna(dms.values[0]):
        dms_code = dms_table["Reference"].values[0]
        if dms_code != header_data.supplier_dms:
            error_message = f"  DMS from database ({dms_code}) does not match DMS from " \
                            f"header ({header_data.supplier_dms}). Check the DMS number and the month number in the " \
                            f"report."
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    else:
        error_message = f"  The DMS reference could not be found for the year {year}, month {month}, and name" \
                        f"{name.irs}. It could be that it is not in the list or that any of these parameters is " \
                        f"written incorrectly. It could not be checked if the DMS number is correct."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    """
    if len(dms_table) > 0:
        dms_code = dms_table["Reference"].values[0]
        if dms_code != header_data.supplier_dms:
            error_message = f"  DMS from database ({dms_code}) does not match DMS from " \
                            f"header ({header_data.supplier_dms}). Check the DMS number and the month number in the " \
                            f"report."
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    else:
        error_message = f"  The DMS reference could not be found for the year {year}, month {month}, and name" \
                        f"{name.irs}. It could be that it is not in the list or that any of these parameters is " \
                        f"written incorrectly. It could not be checked if the DMS number is correct."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    return


def check_report_number_against_kom_date(header_data):
    """Check if the report number is coherent with the months passed since the KoM"""
    # Get KOM date
    day, month, year = [int(x) for x in header_data.kom_date.split(r"/")]
    kom_date = datetime(year, month, day)
    # Get month, year and report number from header
    number, month, year = header_data.report_number.split("_")
    number, month, year = int(number[1:]), int(month[1:]), int(year)
    report_date = datetime(year, month, 1)
    # Check if difference of months is equal to report number
    months_passed = diff_month(report_date, kom_date) + 1
    if months_passed != number:
        error_message = f"  The F4E report number ({header_data.report_number}) and the KOM date " \
                        f"({header_data.kom_date}) have a difference of {months_passed} months"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    return


def check_customer_ref(header_data, name):

    person_reference = f4e_customer_ref.loc[f4e_customer_ref["ESP"].apply(unidecode).isin([name.irs, name.report]), "F4E Customer Reference"].iloc[0]
    if len(person_reference) > 4:
        if header_data.customer_ref != person_reference:
            error_message = f"  The F4E Customer Reference in the report ({header_data.customer_ref}) is different" \
                            f"from the correct reference ({person_reference})"
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    return 


def check_hours_report_vs_header(header_data, document):
    """Check if the number of total hours in the report is the same as the sum of the general activities and the
    specific tasks in the report"""
    general_hours = 0
    specific_hours = 0
    general_taskplan = ""
    specific_taskplans_dic = {}
    # Get the part of the part of the text where the reported hours are
    start = re.search("Main results, achievements and issues encountered during the period:", document).span()[1]
    end = re.search("Main scheduled tasks and milestones for the next period", document).span()[0]
    section = document[start:end]

    # Get the hours done in the period for each task
    while True:
        try:
            match = re.search(r"Task.+:\s*(\d+([,.]\d+)?)\s*hours?\s*", section)
            section = section[match.span()[1]:]
        except AttributeError:
            break
        hours_task = match.group(1).replace(",", ".")
        line = match.group(0)
        try:
            hours_task = float(hours_task)
            if "General Activities".casefold() in line.casefold():
                general_hours += hours_task
                general_taskplan = line[line.find("(") + 1:line.find(")")]
            else:
                specific_hours += hours_task
                specific_taskplans_dic[line[line.find("(") + 1:line.find(")")]] = hours_task
        except ValueError:
            error_message = f"  The number of hours in the line '{line}' could not be transformed to a number." \
                            f"Check if it is written correctly. The hours of that task could not be processed and," \
                            f"probably, the total number of hour will be incorrect because of this."
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    hours_report = general_hours + specific_hours

    # Check if sum of hours in the report is the same as in the header
    if not almost_equal(float(header_data.reported_hours), hours_report):
        error_message = f"  The sum of hours of the tasks ({hours_report}) is not the same as the one found " \
                        f"in the header ({header_data.reported_hours})"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    return general_hours, general_taskplan, specific_taskplans_dic


def check_hours_header_vs_ext_my_time(header_data, name):
    """Check if there are hours in ExtMyTime, the total hours match between ExtMyTime and the report, the general
    activities hours are not more than 8%, the general activities hours in ExtMyTime and the report match and if
    the specific hours in ExtMyTime and the report match."""

    try:
        hours = hours_task_plan[hours_task_plan["Full Name"].apply(unidecode).isin([name.irs_comma, name.irs, name.report])]
        ext_my_time_hours = hours["Total Working hours submitted"].sum()
        general_hours_extmytime = hours[hours["Task Plan Description"].str.contains("General")]["Total Working hours submitted"].values[0]
    except IndexError:
        error_message = f"  Could not find the name '{name.report}' in the list of names file and, consequently, the hours " \
                        f"in the ExtMyTime couldn't be checked"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
        return
    if almost_equal(ext_my_time_hours, 0):
        error_message = "  No hours found in the EXT MY TIME"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
        return
    if not almost_equal(float(header_data.reported_hours), ext_my_time_hours):
        error_message = f"  The total hours as found in the report ({header_data.reported_hours}) don't match the " \
                        f"EXT MY TIME hours ({ext_my_time_hours})"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    general_activities_proportion = general_hours_extmytime / ext_my_time_hours * 100
    if general_activities_proportion > 8:
        error_message = f"  The General Activities task took {float(general_hours_extmytime):.2f} hours, which is a " \
                        f"{float(general_activities_proportion):.2f}% of the total: {ext_my_time_hours} hours"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    return


def check_tasks_hours_report_vs_ext_my_time(header_data, general_hours_report, general_taskplan,
                                            specific_taskplans_dic, hours_task_plan):
    """Check if the hours for each task plan is coincident between the report and ExtMyTime"""
    # Check the general activities task
    try:
        float(general_taskplan)
        general_hours_extmytime = hours_task_plan.loc[hours_task_plan["Task Plan Code"] == int(general_taskplan), "Total Working hours submitted"]
        if not almost_equal(general_hours_extmytime, general_hours_report):
            error_message = f"  The General Activities task hours in ExtMyTime ({float(general_hours_extmytime):.2f}) " \
                            f"is not coincident with the ones declared in the report " \
                            f"({float(general_hours_report):.2f})"
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    except ValueError:
        error_message = f"  The code of the General Activities Task Plan in the report does not match the valid " \
                        f"format: {general_taskplan}. The number of hours in the Task Plan could not be compared " \
                        f"between ExtMyTime and the report."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    # Check the specific activities tasks
    for specific_task in specific_taskplans_dic.keys():
        try:
            float(specific_task)
            specific_hours_extmytime = hours_task_plan.loc[hours_task_plan[hours_task_plan["Task Plan Code"] == int(
                specific_task)].index, "Total Working hours submitted"]
            if len(specific_hours_extmytime) == 0:
                specific_hours_extmytime = 0
            if not almost_equal(specific_hours_extmytime, specific_taskplans_dic[specific_task]):
                error_message = f"  The hours of Specific Task {specific_task} in ExtMyTime " \
                                f"({float(specific_hours_extmytime):.2f}) is not coincident with the ones declared " \
                                f"in the report ({float(specific_taskplans_dic[specific_task]):.2f})"
                print(error_message)
                results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
        except ValueError:
            error_message = f"  The code of a Specific Task Plan in the report does not match the valid format: " \
                            f"{specific_task}. The number of hours in the Task Plan could not be compared between " \
                            f"ExtMyTime and the report."
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    return


def get_codes_activities_section(document, start_text, end_text):
    """
    Gets the numeric codes of the general task and the specific tasks in a section of the text.

    Arguments:
        document (str): whole text of the document
        start_text (str): string that limits the start of the section
        end_text (str): string that limits the end of the section

    Returns:
        str: code of the general task
        list: with the codes of the specific tasks
    """
    general_taskplan_code = ""
    specific_taskplans_codes = []
    # Trim the text to only the wanted part
    start = re.search(start_text, document).span()[1]
    end = re.search(end_text, document).span()[0]
    section = document[start:end]

    while True:
        try:
            match = re.search(r"Task\s.+(F4E-OMF-1159|General).+", section)
            section = section[match.span()[1]:]
        except AttributeError:
            break
        line = match.group(0)
        if "General Activities".casefold() in line.casefold():
            general_taskplan_code = line[line.find("(") + 1:line.find(")")]
        else:
            specific_taskplans_codes = specific_taskplans_codes + [line[line.find("(") + 1:line.find(")")]]

    return general_taskplan_code, specific_taskplans_codes


def check_codes_sections(header_data, general_code, specific_codes, name, section):
    """
    Checks if the codes of tasks in the text are the same as in the Excel file (Hours Task Plan)

    Parameters:
        header_data (class): data read in the header of the document
        folder (str): Name of folder where 'HoursTaskPlan.xlsx' is located
        general_code (str): code of the general task in the document
        specific_codes (list): list of the codes of the specific tasks in the document
        name (class): configurations of the person's name
        section (str): number of section in the document
    """

    hours_person = hours_task_plan[hours_task_plan["Full Name"].apply(unidecode).isin([name.irs_comma, name.irs, name.report])]
    if len(hours_person) == 0:
        error_message = f"  The name of the person {name.irs_comma}, could not be found in the list with the hours in " \
                        f"ExtMyTime. The correspondence of the codes in section 2.2 and 2.4 with the Task Plan " \
                        f"could not be checked."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    else:
        general_code_excel = hours_person[hours_person["Task Plan Description"].str.contains("General Activities")]["Task Plan Code"]
        if general_code not in general_code_excel.values.astype(str):
            error_message = f"  In section {section}, the General Activity code '{general_code}' cannot be found in " \
                            f"the Task Plan Hours. Either the format of the code is not correct or the number of the " \
                            f"activity code is not correct."
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
        specific_codes_excel = hours_person[~hours_person["Task Plan Description"].str.contains("General Activities")]["Task Plan Code"]
        for code in specific_codes:
            if code not in specific_codes_excel.values.astype(str):
                error_message = f"  In section {section}, the Specific Activity code '{code}' cannot be found in the" \
                                f"Task Plan Hours. Either the format of the code is not correct or the number of the" \
                                f"activity code is not correct."
                print(error_message)
                results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    return


def check_filename(filename, header_data):
    """
    Checks if the name of the file follows the structure

    Arguments:
        filename (str): name of the file
    """
    name_split = filename.split()
    if name_split[1] != "Monthly" or name_split[2] != "Report":
        error_message = f"  The name of the file is not according to the template. It should be as follows: " \
                        f"'F4E-OMF-1159-01-01-XX Monthly Report Name Surname #YY MZZ 2023.docx'"
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]


def check_dates_section3(document, header_data):
    # Go to Section 3 part of the text
    start = re.search("Section 3", document).span()[1]
    section = document[start:]

    first_date_start = re.search("Date: ", section).span()[1]
    first_date = section[first_date_start:].split()[0]
    second_date = section[first_date_start:].split()[2]
    if first_date != second_date:
        error_message = f"  The dates in Section 3 are not the same({first_date} vs {second_date})."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    # Get month from report number in header
    month_header = re.match(r"#\d+_M(\d+)_\d+", header_data.report_number).group(1)  # e.g. "3"
    month_header = int(month_header)
    # Get year from report number in header
    year_header = re.match(r"#\d+_M\d+_(\d+)", header_data.report_number).group(1)  # eg."2022"
    year_header = int(year_header)

    date = first_date.split('/')
    month_section3 = int(date[1])
    day_section3 = int(date[0])
    year_section3 = int(date[2])
    if (month_header+1) % 12 != month_section3:
        error_message = f"  The month in Section 3 ({month_section3}) does not correspond to the following month " \
                        f"of the month being reported ({month_header})."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    if day_section3 > 8:
        error_message = f"  The day in Section 3 ({day_section3}) does not correspond to the first 8 days of the month."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]
    if year_section3 != year_header:
        if month_header != 12:
            error_message = f"  The year in Section 3 ({year_section3}) does not correspond to the correct date of " \
                            f"signature of the report."
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    return


def forbidden_words(document, header_data):

    forbidden = ["F4E Project Manager", "F4E Manager", "F4E Line Manager"]
    for word in forbidden:
        if word.lower() in document.lower():
            error_message = f"  The expression '{word}' appears in the body of the document, please delete it."
            print(error_message)
            results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    return


def check_months_header(document, header_data):
    month = header_data.report_number.split('_')[1]

    start = re.search("Tasks and milestones definition for the period", document).span()[0]
    line = document[start:].split('\n')[0]
    if month not in line:
        error_message = f"  The month in the header of Section 2.2 is not valid."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    start = re.search("Main results, achievements and issues encountered during the period", document).span()[0]
    line = document[start:].split('\n')[0]
    if month not in line:
        error_message = f"  The month in the header of Section 2.3 is not valid."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    next_month = int(month[1:]) % 12 + 1
    next_month = f"M{next_month:02}"
    start = re.search("Main scheduled tasks and milestones for the next period", document).span()[0]
    line = document[start:].split('\n')[0]
    if next_month not in line:
        error_message = f"  The month in the header of Section 2.4 is not valid."
        print(error_message)
        results_df.loc[len(results_df)] = [header_data.f4e_reference, name_report, error_message[2:]]

    return


def almost_equal(float_1, float_2):
    return math.isclose(float_1, float_2, rel_tol=0.0001)


def process_monthly(filename, hours_task_plan):
    # Read list of employees
    list_employees = pd.read_excel("LIST OF EMPLOYEES.xlsx")

    global name_report
    print(f"Analyzing {filename.name}...")
    f4e_contract = filename.name.split()[0]
    name_report = ' '.join(filename.name.split()[3:-3])
    document = docx2txt.process(filename)  # Get string of all document
    #document = docx.Document(filename)
    # Get header fields
    header_data = read_header(document)
    # Check if the name of the file follows correct structure
    check_filename(filename.name, header_data)
    # Get the different expressions of the name.
    name = get_names(list_names, filename.name)
    # add_header_hours_to_list(header_data)  # Asked by Arnón, to get a file with all the hours
    # Shows revision number in the header
    show_version_message(header_data)
    # Checks if F4E contracts is the same in the name of the report and the header
    check_f4e_contract(f4e_contract, header_data)
    # Check if DMS in the header and in "DMS Number Monthly Report.xlsx" are the same
    check_supplier_dms(header_data, name, list_employees)
    # Check if number of report (#) is coherent with months passed from KoM
    check_report_number_against_kom_date(header_data)
    # Check if the F4E reference is the same in header and external file
    check_customer_ref(header_data, name)
    # Check if the total number of hours in section 2.3 is the same as in the header
    general_hours_report, general_taskplan, specific_taskplans_dic = check_hours_report_vs_header(header_data, document)
    check_hours_header_vs_ext_my_time(header_data, name)
    # Check if hours for each task plan is the same in the report and ExtMyTime
    check_tasks_hours_report_vs_ext_my_time(header_data, general_hours_report, general_taskplan, specific_taskplans_dic, hours_task_plan)
    # Get numerical code of tasks in sections 2.2 and 2.4
    general_code_22, specific_codes_22 = get_codes_activities_section(document, "definition for the period:", "during the period")
    general_code_24, specific_codes_24 = get_codes_activities_section(document, "for the next period:", "List of documents")
    # Check numerical Codes of tasks in sections 2.2 and 2.4
    check_codes_sections(header_data, general_code_22, specific_codes_22, name, "2.2")
    check_codes_sections(header_data, general_code_24, specific_codes_24, name, "2.4")
    # Check both dates in section 3 are the same
    check_dates_section3(document, header_data)
    # Check there are no "forbidden words" in the text
    forbidden_words(document, header_data)

    #Check months headers
    check_months_header(document, header_data)
    return



