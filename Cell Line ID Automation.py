# Author: Kenny Ma
# Contact: 626-246-2233 or kyoma17@gmail.com
Date =  "2023-July-6"
version = 3.6
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
import tkinter as tk
import tkinter
from tkinter import messagebox
import pandas
import pandas as pd
import warnings
import os

warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

from utils.SampleSelector import selectSample
from utils.WordConsolidator import consolidateWordOutputs
from utils.SampleProcessor import processSamples

########################################################################################################################
def main():
    print("Welcome to the ClimaSTR Cell Line ID script" '\n' "Please select the input file from the GMapper Program")
    print("version: " + str(version))
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
    if not os.path.exists("CellLineTEMP"):
        os.makedirs("CellLineTEMP")

    # read .xlsx into pandas dataframe
    df = pandas.read_excel(file_path)
    print("Loaded file: " + file_path + '\n')

    # Runs the Selenium 
    result_collection, sample_order = processSamples(df)

    clear_temp_folder()

    # Bulk select the samples
    for each in result_collection:
        results = each[0]
        sampleName = each[1]

        selectSample(results, sampleName)


    # run vba macro to save all word documents
    consolidateWordOutputs(sample_order, selected_client_info, reference_number)

    # Script is finished
    print("Script is finished. Please check the output folder for the results.")
    quit()



########################################################################################################################
# GUI Functions

def SelectClient():
    # Selects the client from the listbox and returns the order number
    print("Select Client from the listbox and enter the order number")

    # Load Client Data from Excel File and create a dataframe
    client_database = pd.read_excel("CellLineClients.xlsx")
    client_list = client_database["Nickname"].tolist()

    selected_item = ""
    order_number = ""

    def submit():
        # Get the selected client and order number
        nonlocal selected_item
        nonlocal order_number

        # Window Title "Please Select a Client"
        
        selected_item = listbox.get(listbox.curselection())
        order_number = order_entry.get()
        
        print("Selected Client:", selected_item)
        print("Order number:", order_number)

        # Close the window
        root.quit()
        root.destroy()

    root = tk.Tk()
    root.title("Order Form")
    root.geometry("300x300")

    # Create a listbox with the client names and a submit button
    label = tk.Label(root, text="Please Select a Client")


    listbox = tk.Listbox(root)
    for item in client_list:
        listbox.insert(tk.END, item)

    label.pack()
    listbox.pack()

    order_label = tk.Label(root, text="Order Number:")
    order_label.pack()

    order_entry = tk.Entry(root)
    order_entry.pack()

    submit_button = tk.Button(root, text="Submit", command=submit)
    submit_button.pack()

    root.mainloop()

    return selected_item, order_number

def display_readme():
    try:
        with open('readme.txt', 'r') as file:
            readme_content = file.read()
            messagebox.showinfo("Read Me", readme_content)
    except FileNotFoundError:
        messagebox.showerror("Error", "readme.txt not found.")

    # Create the main Tkinter window
    root = tk.Tk()

    # Set window title and size
    root.title("My Program")
    root.geometry("300x200")

    # Create a button to show the Read Me message
    readme_button = tk.Button(root, text="Read Me", command=display_readme)
    readme_button.pack(pady=50)

    # Start the Tkinter event loop
    root.mainloop()

def show_done_window():
    def show_done_message():
        messagebox.showinfo("Done", "Cell Line ID script has finished running!")
        root.destroy()  # Close the "Done" window and exit the program

    # Create the main Tkinter window
    root = tk.Tk()

    # Set window title and size
    root.title("Done")
    root.geometry("300x200")

    # Create a label to display the "Done" message
    message_label = tk.Label(root, text="Process completed!")
    message_label.pack(pady=50)

    # Create a button to close the window and exit the program
    done_button = tk.Button(root, text="Done", command=show_done_message)
    done_button.pack()

    # Start the Tkinter event loop
    root.mainloop()

########################################################################################################################
# Helper Functions
def clear_temp_folder():
    # Delete all files in the CellLineTEMP folder
    folder = "CellLineTEMP"
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                print("Deleting " + filename)
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                print("Deleting " + filename)
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))




if __name__ == "__main__":
    main()