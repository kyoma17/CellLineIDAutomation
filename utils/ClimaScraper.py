from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.firefox.options import Options

import pandas
import time
from bs4 import BeautifulSoup
import threading
from utils.DictionaryGenerator import generateReplacementDictionary
from params import waitTime


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
    options.headless = True
    
    driver = webdriver.Firefox(options=options)

    # Open a web browser and navigate to the website
    driver.get('http://bioinformatics.hsanmartino.it/clima2/index.php')
    time.sleep(waitTime)

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
    time.sleep(waitTime)

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

    # export data to csv
    try:
        tableDF = pandas.DataFrame(data, columns=column_names)
    except ValueError:
        # print("There may be an error with sample " + sampleName + " from Clima")

        # The data may not have the same number of columns as the column names
        # Insert empty columns to make the number of columns match the number of column names
        
        data = [row + [""] * (len(column_names) - len(row)) for row in data]
        tableDF = pandas.DataFrame(data, columns=column_names)

    # If no results are found, return an empty dataframe
    if tableDF.empty:
        print("No results found for sample " + sampleName + " from Clima")



    # PandasTableDF = pandas.read_html(html_table)[0]
    # PandasTableDF.reset_index(drop=True, inplace=True)

    # try:
    #     tableDF.columns = PandasTableDF.columns.get_level_values(0)
    # except ValueError:
    #     PandasTableDF.to_csv("ClimaError.csv")
    #     tableDF.to_csv("ClimaError2.csv")
    #     raise ValueError("Clima Error")


    driver.quit()

    # number of results to return
    # check how many rows in the table
    # numberOfHits = len(tableDF.index) - 2

    # if numberOfHits < 10:
    #     number_of_results = numberOfHits
    # else:
    #     number_of_results = 10

    # def get_best_match(tableDF, i):
    #     # print how many rows in the table
    #     bestMatched.append(tableDF.iloc[i])

    bestMatched = []
    threads = []
    # for i in range(2, number_of_results + 2):

    #     thread = threading.Thread(target=get_best_match, args=(tableDF, i))
    #     threads.append(thread)
    #     thread.start()
    
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