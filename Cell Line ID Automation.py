# Author: Kenny Ma
# Contact: 626-246-2233 or kyoma17@gmail.com
Date =  "2023-September-8"
version = "3.7"
# This is the Main Operating Script for the Cell Line ID Automation Program

# Description: This script will take in an excel file from the GMapper program and perform the ClimaSTR Cell Line ID script on each sample
# The script will then consolidate the results into a single .docx file
# Requirements: Selenium, Pandas, BeautifulSoup, docx, tkinter, Firefox, geckodriver.exe
# there is a Requirements.txt file that can be used to install the required modules
# Use pip install -r Requirements.txt to install the required modules
# GekoDriver: Place geckodriver.exe in the same directory as this script
# Instructions: Run the script and select the input file from the GMapper program. The script will then perform the ClimaSTR Cell Line ID script on each sample and output a .docx file for each sample.
# The input file must have the following information: Sample Name, D5S818, D13S317, D7S820, D16S539, vWA, TH01, AMEL, TPOX, CSF1PO, D21S11

# You must have Word installed on your computer with the Trust Center settings set to "Enable all macros"

from tkinter import filedialog
import tkinter
import pandas
import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

from utils.SampleSelector import selectSample
from utils.WordConsolidator import consolidateWordOutputs
from utils.SampleProcessor import processSamples
from utils.SelectClient import SelectClient
from utils.PrepTemp import PrepTempFolder
from utils.TemplateWriter import fillTemplate
from params import debug

########################################################################################################################
def main():
    print("Welcome to the ClimaSTR Cell Line ID script" '\n' "Please select the input file from the GMapper Program")
    print("version: " + str(version))

    if debug:
        print("WARNING: Debug mode is enabled. Best Result Selection will not Show Up")
        print("Highest scoring result will be selected automatically")
        file_path = "TestFiles/20230906 Wood BMS-all.xlsx"
        # file_path = "TestFiles/DebugOL2.xlsx"
        # file_path = "TestFiles/GP10Single.xlsx"
        # file_path = "TestFiles/SuperTest.xlsx"
        # file_path = "TestFiles/SuperTest - Double.xlsx"
        selected_client = "Wood"
        reference_number = "testtest"
    else:
        # get file path from user using tkinter file dialog
        tkinter.Tk().withdraw()
        # open file dialog in current directory and allow excel files only
        file_path = filedialog.askopenfilename(initialdir=".", title="Select file", filetypes=(
            ("Select CellLine ID file", "*.xlsx"), ("all files", "*.*")))

        # Load the client list excel file
        selected_client, reference_number = SelectClient()

    client_database = pd.read_excel("CellLineClients.xlsx")

    # Get the client row from the client database
    selected_client_info = client_database.loc[client_database["Nickname"] == selected_client]

    # If the output folder does not exist, create it
    PrepTempFolder()

    # Load .xlsx into pandas dataframe
    df = pandas.read_excel(file_path)
    print("Loaded file: " + file_path + '\n')

    # Runs the Selenium script on each sample
    result_collection, sample_order = processSamples(df)

    selected_results = []

    # Bulk select the samples in the sample order
    for each_sample in sample_order:
        # locate the sample in the results collection
        for each in result_collection:
            results = each[0]
            sampleName = each[1]
            if sampleName == each_sample:
                selected_results.append(selectSample(results, sampleName))
                break


    # run vba macro to save all word documents
    consolidateWordOutputs(sample_order, selected_client_info, reference_number)

    # Script is finished
    print("Script is finished. Please check the output folder for the results.")
    # input("Press Enter to exit...")
    quit()

########################################################################################################################
if __name__ == "__main__":
    main()