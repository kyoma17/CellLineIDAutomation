import docx
import time
import win32com.client
import os

def consolidateWordOutputs(listOfSamples, clientInfo, reference_number):
    # Create Replace Dictionary for the final report
    # This dictionary is for the Header in the Paperwork
    replaceDict = {"_PIName": clientInfo["PIName"].values[0],
                   "_ClientName": clientInfo["ClientName"].values[0],
                   "_ClientEmail": clientInfo["ClientEmail"].values[0],
                   "_ClientPhoneNumber": clientInfo["ClientPhoneNumber"].values[0],
                   "_Institution": clientInfo["Institution"].values[0],
                   "_ReferenceNumber": reference_number,
                   "_Date": time.strftime("%m/%d/%Y"),
                   "_SampleNumber": len(listOfSamples),
                   "_Batches": "1",
                   }


    # Open the Header and Footer template
    combined_document = docx.Document('ClientTemplate.docx')

    # iterate through the other documents and add them to the combined document

    # add the content of each document to the combined document keep the formatting each document
    # add a page break to the end of each document except the last one
    for doc in listOfSamples:
        temp_doc = docx.Document("CellLineTEMP/" + doc + ".docx")
        for element in temp_doc.element.body:
            combined_document.element.body.append(element)
        # combined_document.add_page_break()

    # Show header and footer from first page in all pages
    for section in combined_document.sections:
        section.different_first_page_header_footer = False

    # Replace spaces in client name with underscores
    clientFileName = clientInfo["Nickname"].values[0].replace(" ", "_")

    # Save the final report in CellLineOutput folder
    report_name = "CellLineOutput/" + clientFileName + "_Cell_Line_ID_" + reference_number + ".docx" 
    combined_document.save(report_name)

    # Convert the Python dictionary to a VBA dictionary
    vba_dict = win32com.client.Dispatch("Scripting.Dictionary")

    # Iterate through the Python dictionary and add the key/value pairs to the VBA dictionary
    for key in replaceDict:
        vba_dict.Add(key, replaceDict[key])

    # get the absolute path of the word document
    report_name = os.path.abspath(report_name)

    # open word document
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = 0
    print(report_name)
    doc = word.Documents.Open(report_name, ReadOnly=1)

    # Load the Header Editor VBA code from the file
    with open("MSO/HeaderEditor.bas", "r") as f:
        vbaCode = f.read()

    # Load the Page Breaker VBA code from the file
    with open("MSO/PageBreaker.bas", "r") as f:
        vbaCode += f.read()


    # Inject vba script into word document
    doc.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vbaCode)

    # run macro
    word.Run("ReplaceHeaderKeyword", vba_dict)
    word.Run("AddPageBreaks")

    # save and close word document
    doc.Save()
    doc.Close()

