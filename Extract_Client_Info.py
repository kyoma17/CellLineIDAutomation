# Load all the word documents in current directory, extract the client name and save it in a csv file
# Information of interested is located in a table in the header of the document
# The table is in the format of 2 columns, the first column is the name of the information, the second column is the value of the information
# We are interested in the information of "Pi Name", "Institution", "Client Name",  "Client Email" and "Client Phone Number"

# PI Name:	Supreet Agarwal	Institution:	NIH
# Client Name:	Supreet Agarwal	    Order Number:	083022
# Client Email:	Suprett.agarwal@nih.gov
# Batch(es):	083422
# Client Phone Number:	(240)-760-7099	Number of Samples:	34

# Create a dataframe to store the information
# Save the dataframe to a csv file

import os
import pandas as pd
import docx

# Get the current working directory
cwd = os.getcwd()

# Get all the word documents in the current directory, ignore temporary files
files = [f for f in os.listdir(cwd) if f.endswith('.docx') and not f.startswith('~$')]

# print(files)

# Create a dataframe to store the information
df = pd.DataFrame(columns=['Pi Name', 'Institution', 'Client Name', "Client Email", "Client Phone Number"])

# Loop through all the word documents
for file in files:
    # Load the word document
    doc = docx.Document(file)

    # Print the content of the header   
    
    # Get the table in the header
    table = doc.sections[0].header.tables[0]

    # Get the information of "Pi Name", "Institution", "Client Name",  "Client Email" and "Client Phone Number"
    pi_name = table.cell(2,1).text
    client_name = table.cell(3,1).text
    client_email = table.cell(4,1).text
    client_phone = table.cell(5,1).text
    institution = table.cell(2,3).text


    # Add the information to the dataframe
    df = df.append({'Pi Name': pi_name, 'Institution': institution, 'Client Name': client_name, "Client Email": client_email, "Client Phone Number": client_phone}, ignore_index=True)

    # Save the dataframe to a csv file
    df.to_csv('Client_Info.csv', index=False)
