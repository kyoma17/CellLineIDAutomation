from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.firefox.options import Options

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas
import time
from bs4 import BeautifulSoup
import threading
from utils.DictionaryGenerator import generateReplacementDictionary
from params import waitTime, Headless, debug


def ClimaSTRSearch(sampleName, sampleDF, sampleNumber):
    # Takes in a pandas dataframe of a single sample and performs the selenium script against Clima
    # Will return a pandas dataframe of the results of web scraping

    # Webpage input fields
    D5S818_list = ["D5S818", "D5S818_data1","D5S818_data2", "D5S818_data3", "D5S818_data4"]
    D13S317_list = ["D13S317", "D13S317_data1","D13S317_data2", "D13S317_data3", "D13S317_data4"]
    D7S820_list = ["D7S820", "D7S820_data1","D7S820_data2", "D7S820_data3", "D7S820_data4"]
    D16S539_list = ["D16S539", "D16S539_data1","D16S539_data2", "D16S539_data3", "D16S539_data4"]
    VWA_list = ["vWA", "VWA_data1", "VWA_data2", "VWA_data3", "VWA_data4"]
    TH01_list = ["TH01", "TH01_data1","TH01_data2", "TH01_data3", "TH01_data4"]
    Amelogenin_list = ["AMEL", "AMG_data1", "AMG_data2"]
    TPOX_list = ["TPOX", "TPOX_data1","TPOX_data2", "TPOX_data3", "TPOX_data4"]
    CSF1PO_list = ["CSF1PO", "CSF1PO_data1", "CSF1PO_data2","CSF1PO_data3", "CSF1PO_data4"]
    # D21S11_list = ["D21S11", "D21S11_data1", "D21S11_data2", "D21S11_data3", "D21S11_data4"]

    master_list = [D5S818_list, D13S317_list, D7S820_list, D16S539_list,
                   VWA_list, TH01_list, Amelogenin_list, TPOX_list, CSF1PO_list]

    # Create a new instance of the Firefox driver
    options = Options()
    options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    webdriver_path = r"/webdrivers/geckodriver.exe"
    options.headless = Headless
    
    driver = webdriver.Firefox(options=options)

    # Open a web browser and navigate to the website
    driver.get('http://bioinformatics.hsanmartino.it/clima2/index.php')
    
    # Wait for the page to load
    WebDriverWait(driver, waitTime).until(EC.presence_of_element_located((By.ID, "usr_email")))


    print("Collecting data for sample " + sampleName + " from Clima...")

    # Input each allele into the webpage from the dataframe
    for each in master_list:
        marker = each[0]
        row = sampleDF.loc[sampleDF["Markers"] == marker]

        for i in range(1, len(each)):
            # skip nan values
            if pandas.isna(row["Allele" + str(i)].values[0]):
                continue

            allele = driver.find_element("id", each[i])

            allele.send_keys(row["Allele" + str(i)].values[0])

    # enter email and country and submit
    email = driver.find_element("id", "usr_email")
    email.send_keys("info@laragen.com")
    country = driver.find_element("id", "usr_country")
    country.send_keys("United States")

    # Find the submit button
    submit_button = driver.find_element(
        By.XPATH, "//input[@type='submit' and @value='submit']")


    # Click the submit button
    submit_button.click()

    # Wait for the page to load
    WebDriverWait(driver, waitTime).until(EC.presence_of_element_located((By.XPATH, "(//table)[3]")))
    

    # retreive table from webpage and convert to pandas dataframe
    table = driver.find_element(By.XPATH, "(//table)[3]")
    html_table = table.get_attribute("outerHTML")
    soup = BeautifulSoup(html_table, "html.parser")
    souptable = soup.find("table")
    rows = souptable.find_all("tr")
    header_row = rows[0]
    column_names = [th.text for th in header_row.find_all("th")]


    # Store the data in a list of lists
    data = []
    for row in rows:
        cells = row.find_all("td")
        values = [cell.text for cell in cells]
        data.append(values)

    number_rows = len(data)
    column_length = len(column_names)
    data_columnLength = len(data[0])
    try:

        if number_rows <= 2:
            print("No results found for sample " + sampleName)
            data_dict = {
                "Name": sampleName,
                "Dataset": "Clima",
                "Cat. No.": "N/A",
                "CVCL": "N/A",
                "% Match": "0%",  
                "D5S818": "",
                "D13S317": "",
                "D7S820": "",
                "D16S539": "",
                "VWA": "",
                "TH01": "",
                "AMG": "",
                "TPOX": "",
                "CSF1PO": "",
                "D18S51": "",
                "D19S433": "",
                "D21S11": "",
                "D2S1338": "",
                "D3S1358": "",
                "D8S1179": "",
                "FGA": "",
                "PentaD": "",
                "PentaE": ""
            }
            # create a pandas dataframe from the data
            tableDF = pandas.DataFrame(data_dict, index=[2])
            #fill index 0 and 1 with None
            tableDF.loc[0] = "None"
            tableDF.loc[1] = "None"
            #Sort the index
            tableDF = tableDF.sort_index()

        elif data_columnLength < column_length:
            # print("Column match issue")
            # This occures when the Webpage does not show the last columns
            try :
                tableDF = pandas.DataFrame(data, columns=column_names)

            except:
                # Add empty columns to the data
                for i in range(column_length - data_columnLength):
                    data[1].append("")

                # create a pandas dataframe from the data
                tableDF = pandas.DataFrame(data, columns=column_names)
                tableDF.loc[1] = "None"
                tableDF = tableDF.sort_index()


        else:
            tableDF = pandas.DataFrame(data, columns=column_names)
    
    except:
        print("Error with tableDF")
        print(data)
        print(column_names)
        print(data_columnLength)
        print(column_length)
        print(sampleName)
        sampleDF.to_csv("Debug/" + sampleName + ".csv")

    if debug:
        print("Clima Results for " + sampleName + ":")
        print(tableDF)


    driver.quit()

    bestMatched = []
    threads = []
    
    number_of_results = 10

    # Create a thread for the first 10, if there are 10, otherwise create a thread for each result
    for i in range(2, number_of_results + 2):
        try :
            thread = threading.Thread(target=bestMatched.append, args=(tableDF.iloc[i],))
            threads.append(thread)
            thread.start()
        except IndexError:
            # go to next result
            break
    # Wait for all threads to complete
    for thread in threads:
        thread.join()

    replacementsDictionaries = []

    # Create a list to store the threads
    threads = []

    # Create a function to run in a separate thread
    def generate_replacement_dictionary_thread(sampleName, sampleDF, bestMatched, prefix):
        replacementsDictionaries.append(generateReplacementDictionary(
            sampleName, sampleDF, bestMatched, prefix, sampleNumber))

    # Create a thread for each call to generateReplacementDictionary
    for i in range(len(bestMatched)):
        t = threading.Thread(target=generate_replacement_dictionary_thread, args=(
            sampleName, sampleDF, bestMatched[i], "Clima"))
        threads.append(t)

    # Start the threads
    for t in threads:
        t.start()

    # Wait for the threads to finish
    for t in threads:
        t.join()

    return replacementsDictionaries