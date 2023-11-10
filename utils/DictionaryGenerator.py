def generateReplacementDictionary(sampleName, sampleDF, bestMatched, website, sampleNumber):
    '''
    Generates a dictionary of the replacement values for the template
    The replacement values are taken from the sample dataframe and the best matched result
    to be used to replace the values in the template
    Args:
        sampleName: The name of the sample
        sampleDF: The dataframe of the sample
        bestMatched: The best matched result
        website: The website the sample was run on
        sampleNumber: The sample number
    Returns:
        replacementsDictionary: The dictionary of the replacement values
    '''
    
    # Clima dictionary for the template
    if website == "Clima":
        replacementsDictionary = {
            # Main Info
            "_SAMPLE_NAME": sampleName,
            "_sampleNumber": sampleNumber,
            "website": "Clima2",
            "test": sampleDF["Test Name"],

            # Data from the Results highest scoring match
            "_dataset": bestMatched["Dataset"],
            "_bMatchScore": bestMatched["% Match"],
            "_bMatchName": bestMatched["Name"],
            "_bMatchCellLineNo": bestMatched["Cat. No."],

            "D5S818_bM": bestMatched["D5S818"],  # Marker 1
            "D13S317_bM": bestMatched["D13S317"],  # Marker 2
            "D7S820_bM": bestMatched["D7S820"],  # Marker 3
            "D16S539_bM": bestMatched["D16S539"],  # Marker 4
            "vWA_bM": bestMatched["VWA"],  # Marker 5
            "TH01_bM": bestMatched["TH01"],  # Marker 6
            "AMEL_bM": bestMatched["AMG"],  # Marker 7
            "TPOX_bM": bestMatched["TPOX"],  # Marker 8
            "CSF1PO_bM": bestMatched["CSF1PO"],  # Marker 9
            "D21S11_bM": bestMatched["D21S11"],  # Marker 10
        }
        if sampleDF["Test Name"].values[0] == "GenePrint_24_POP7_Panels_v1.0":
            replacementDictionaryGP24 = {
                "D10S1248_bM": "",  # Marker 11
                "D12S391_bM": "",  # Marker 12
                "D18S51_bM": bestMatched["D18S51"],  # Marker 13
                "D19S433_bM": bestMatched["D19S433"],  # Marker 14
                "D1S1656_bM": "",  # Marker 15
                "D22S1045_bM": "",  # Marker 16
                "D2S1338_bM": bestMatched["D2S1338"],  # Marker 17
                "D2S441_bM": "",  # Marker 18
                "D3S1358_bM": bestMatched["D3S1358"],  # Marker 19
                "D8S1179_bM": bestMatched["D8S1179"],  # Marker 20
                "DYS391_bM": "",  # Marker 21
                "FGA_bM": bestMatched["FGA"],  # Marker 22
                "Penta E_bM": bestMatched["PentaD"],  # Marker 23
                "Penta D_bM": bestMatched["PentaE"],  # Marker 24
            }
            replacementsDictionary.update(replacementDictionaryGP24)

    # Expasy dictionary for the template
    elif website == "Expasy":
        replacementsDictionary = {

            "_SAMPLE_NAME": sampleName,
            "website": "Expasy",
            "test": sampleDF["Test Name"],
            "_sampleNumber": sampleNumber,

            # Data from the Results highest scoring match
            "_dataset": "Expasy",
            "_bMatchScore": bestMatched["Score"],
            "_bMatchName": bestMatched["Name"],
            "_bMatchCellLineNo": bestMatched["Accession"],

            "D5S818_bM": bestMatched["D5S818"],  # Marker 1
            "D13S317_bM": bestMatched["D13S317"],  # Marker 2
            "D7S820_bM": bestMatched["D7S820"],  # Marker 3
            "D16S539_bM": bestMatched["D16S539"],  # Marker 4
            "vWA_bM": bestMatched["vWA"],  # Marker 5
            "TH01_bM": bestMatched["TH01"],  # Marker 6
            "AMEL_bM": bestMatched["Amel"],  # Marker 7
            "TPOX_bM": bestMatched["TPOX"],  # Marker 8
            "CSF1PO_bM": bestMatched["CSF1PO"],  # Marker 9
            "D21S11_bM": bestMatched["D21S11"],  # Marker 10
        }

        # If GenePrint24 Add the extra markers
        if sampleDF["Test Name"].values[0] == "GenePrint_24_POP7_Panels_v1.0":
            replacementDictionaryGP24 = {
                "D10S1248_bM": bestMatched["D10S1248"],  # Marker 11
                "D12S391_bM": bestMatched["D12S391"],  # Marker 12
                "D18S51_bM": bestMatched["D18S51"],  # Marker 13
                "D19S433_bM": bestMatched["D19S433"],  # Marker 14
                "D1S1656_bM": bestMatched["D1S1656"],  # Marker 15
                "D22S1045_bM": bestMatched["D22S1045"],  # Marker 16
                "D2S1338_bM": bestMatched["D2S1338"],  # Marker 17
                "D2S441_bM": bestMatched["D2S441"],  # Marker 18
                "D3S1358_bM": bestMatched["D3S1358"],  # Marker 19
                "D8S1179_bM": bestMatched["D8S1179"],  # Marker 20
                "DYS391_bM": bestMatched["DYS391"],  # Marker 21
                "FGA_bM": bestMatched["FGA"],  # Marker 22
                "Penta E_bM": bestMatched["Penta E"],  # Marker 23
                "Penta D_bM": bestMatched["Penta D"],  # Marker 24
            }
            replacementsDictionary.update(replacementDictionaryGP24)

    # Replacement Dictionary for the Template from  Data Input
    markers = ["D5S818", "D13S317", "D7S820", "D16S539",
               "vWA", "TH01", "AMEL", "TPOX", "CSF1PO", "D21S11"]

    # If GenePrint24 Add the extra markers
    if sampleDF["Test Name"].values[0] == "GenePrint_24_POP7_Panels_v1.0":
        markers.extend(["D10S1248", "D12S391", "D18S51", "D19S433", "D1S1656", "D22S1045",
                        "D2S1338", "D2S441", "D3S1358", "D8S1179", "DYS391", "FGA", "Penta E", "Penta D"])

    allele_columns = ["Allele1", "Allele2", "Allele3", "Allele4"]

    template_dict = {}

    for marker in markers:
        for i, allele_column in enumerate(allele_columns, start=1):
            key = f"{marker}_{i}"

            value = sampleDF.loc[sampleDF["Markers"]
                                 == marker][allele_column].values[0]
            template_dict[key] = value

    replacementsDictionary.update(template_dict)

    return replacementsDictionary
