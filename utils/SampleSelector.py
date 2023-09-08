import tkinter as tk
from tkinter import ttk
import docx
import pandas
import win32com.client
import os
import psutil
from utils.TemplateWriter import fillTemplate

from params import debug


def selectSample(bestMatchedSamples, sampleName):
    # Display the best matched samples to the user and ask for user input
    window = tk.Tk()
    window.title("Select Best Result for " + sampleName)
    window.columnconfigure(0, minsize=250, weight=1)
    window.rowconfigure([0, 1], minsize=200, weight=1)

    # # Center window in the left of the screen
    # windowWidth = window.winfo_reqwidth()
    # positionRight = int(window.winfo_screenwidth() / 2 - windowWidth / 2)
    # positionDown = int(window.winfo_screenheight() / 2 - windowHeight / 2)
    # window.geometry("+{}+{}".format(positionRight, positionDown))
    # window.geometry("{}x{}".format(windowWidth, windowHeight))

    tree = ttk.Treeview(window)

    columns = ("name", "_dataset", "_bMatchScore", "_bMatchName", "_bMatchCellLineNo", "D5S818_bM", "D13S317_bM",
               "D7S820_bM", "D16S539_bM", "vWA_bM", "TH01_bM", "AMEL_bM", "TPOX_bM", "CSF1PO_bM", "D21S11_bM")
    
    # Create the treeview
    tree = ttk.Treeview(window, columns=columns, show='headings')

    # Set the column widths
    for col in columns:
        tree.column(col, width=100, anchor=tk.W)

    # Set the column headings
    for col, heading in enumerate(columns):
        tree.heading(col, text=heading, anchor=tk.W)

    treeviewDictionary = {}

    climaCounter = 1
    expasyCounter = 1

    # If there are no results for both websites, enter "No Results" into the bestMatchedSamples list
    # if len(bestMatchedSamples) == 0:
    #     NR = "No Results"
    #     bestMatchedSamples.append({"website": "Clima2", "_dataset": NR, "_bMatchScore": NR, "_bMatchName": NR, "_bMatchCellLineNo": NR, "D5S818_bM": NR, "D13S317_bM": NR,
    #            "D7S820_bM": NR, "D16S539_bM": NR, "vWA_bM": NR, "TH01_bM": NR, "AMEL_bM": NR, "TPOX_bM": NR, "CSF1PO_bM": NR, "D21S11_bM": NR})
    #     bestMatchedSamples.append({"website": "Expasy", "_dataset": NR, "_bMatchScore": NR, "_bMatchName": NR, "_bMatchCellLineNo": NR, "D5S818_bM": NR, "D13S317_bM": NR,
    #             "D7S820_bM": NR, "D16S539_bM": NR, "vWA_bM": NR, "TH01_bM": NR, "AMEL_bM": NR, "TPOX_bM": NR, "CSF1PO_bM": NR, "D21S11_bM": NR})
        
    for i, data in enumerate(bestMatchedSamples):
        if data["website"] == "Clima2":
            treeviewDictionary[f"Clima{climaCounter}"] = data
            climaCounter += 1
        else:
            treeviewDictionary[f"Expasy{expasyCounter}"] = data
            expasyCounter += 1

    for i, (name, data) in enumerate(treeviewDictionary.items()):
        tree.insert("", i, text=name, values=(name, data["_dataset"], data["_bMatchScore"], data["_bMatchName"], data["_bMatchCellLineNo"], data["D5S818_bM"],
                    data["D13S317_bM"], data["D7S820_bM"], data["D16S539_bM"], data["vWA_bM"], data["TH01_bM"], data["AMEL_bM"], data["TPOX_bM"], data["CSF1PO_bM"], data["D21S11_bM"]))


    selected_result = None

# Create a submit button, close window when clicked
    def submit():
        # Get the selected item
        item = tree.selection()[0]
        # Print the Name of the selected item
        selection = tree.item(item, "values")[0]
        print("Selected result: ", selection, " with ", treeviewDictionary[selection]["_bMatchScore"], " match score")
        # Pull selected item from dictionary and write to template
        fillTemplate(treeviewDictionary[selection])

        selected_result = treeviewDictionary[selection]

        # Close the window
        window.destroy()
        window.quit()
       
    if debug:
        # Skip the GUI and just select the Highest Match Score
        # Find the highest match score
        highestMatchScore = 0
        highestMatchScoreIndex = 0
        for i, data in enumerate(bestMatchedSamples):
            matchScore = float(data["_bMatchScore"].replace("%", ""))
            if matchScore > highestMatchScore:
                highestMatchScore = matchScore
                highestMatchScoreIndex = i

        # Get the selected result with the highest match score
        try:
            selected_result = bestMatchedSamples[highestMatchScoreIndex]
        except:
            print("No results found for sample " + sampleName)
            print(bestMatchedSamples)
            print(highestMatchScoreIndex)
            print(len(bestMatchedSamples))
            print("Exiting...")

        selection_name = selected_result["_bMatchName"]
        selection_score = selected_result["_bMatchScore"]

        # Print the selected result and its match score
        print("Selected result: ", selection_name, " with ", selection_score, " match score")
        # Pull selected item from dictionary and write to template
        fillTemplate(selected_result)

        # Close the window
        window.destroy()
        window.quit()
        return 

    submit_button = tk.Button(window, text='SELECT RESULT', command=submit)
    tree.pack()
    submit_button.pack()

    # Run the main loop
    window.mainloop()

########################################################################################################################
