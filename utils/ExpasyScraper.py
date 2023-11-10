from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.firefox.options import Options

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas
from bs4 import BeautifulSoup
import threading
from utils.DictionaryGenerator import generateReplacementDictionary
from params import waitTime, Headless, debug
import time


def retry_scraper(scraper_function, max_retries=3, delay=2):
    '''
    A decorator that wraps the passed in function and retries
    '''
    def wrapper(*args, **kwargs):
        attempts = 0
        while attempts < max_retries:
            try:
                return scraper_function(*args, **kwargs)
            except Exception as e:
                attempts += 1
                print(f"Retry {attempts}/{max_retries} for {scraper_function.__name__} due to error: {e}")
                time.sleep(delay)
        print(f"All retries failed for {scraper_function.__name__}")
        return None  # Or handle the failure as needed
    return wrapper


@retry_scraper
def ExpasySTRSearch(sampleName, sampleDF, sampleNumber):
    '''
    Takes in a pandas dataframe of a single sample and performs the selenium script against Expasy
    Will return a pandas dataframe of the results of web scraping
    Args:
        sampleName: The name of the sample
        sampleDF: The dataframe of the sample
        sampleNumber: The sample number
    '''
    # Webpage input fields
    Amelogen = ["AMEL", "input-Amelogenin"]
    CSF1PO = ["CSF1PO", "input-CSF1PO"]
    D2S133 = ["D2S133", "input-D2S133"]
    D3S135 = ["D3S1358", "input-D3S1358"]
    D5S818 = ["D5S818", "input-D5S818"]
    D7S820 = ["D7S820", "input-D7S820"]
    D8S1179 = ["D8S1179", "input-D8S1179"]
    D13S317 = ["D13S317", "input-D13S317"]
    D16S539 = ["D16S539", "input-D16S539"]
    D18S51 = ["D18S51", "input-D18S51"]
    D19S433 = ["D19S433", "input-D19S433"]
    D21S11 = ["D21S11", "input-D21S11"]
    FGA = ["FGA", "input-FGA"]
    PentaD = ["Penta D", "input-Penta_D"]
    PentaE = ["Penta E", "input-Penta_E"]
    TH01 = ["TH01", "input-TH01"]
    TPOX = ["TPOX", "input-TPOX"]
    vWA = ["vWA", "input-vWA"]
    D2S1338 = ["D2S1338", "input-D2S1338"]

    D10S1248 = ["D10S1248", "input-D10S1248"]
    D12S391 = ["D12S391", "input-D12S391"]
    D22S1045 = ["D22S1045", "input-D22S1045"]
    D2S441 = ["D2S441", "input-D2S441"]
    D1S1656 = ["D1S1656", "input-D1S1656"]
    DYS391 = ["DYS391", "input-DYS391"]

    master_list = [Amelogen, CSF1PO, D2S133, D3S135, D5S818, D7S820, D8S1179,
                   D13S317, D16S539, D18S51, D19S433, D21S11, FGA, PentaD, PentaE,
                   TH01, TPOX, vWA, D10S1248, D12S391, D22S1045, D2S441, D1S1656, DYS391, D2S1338]

    # Create a new instance of the Firefox driver
    options = Options()
    options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    options.headless = Headless

    driver = webdriver.Firefox(options=options)

    # go to the Expasy STR website
    driver.get("https://www.cellosaurus.org/str-search/")

    # Wait for the page to load
    WebDriverWait(driver, waitTime).until(EC.element_to_be_clickable((By.ID, "search")))

    print("Collecting data for sample " + sampleName + " from Expasy...")
    # Click Checkboxes if GP24
    if "GenePrint_24_POP7_Panels_v1.0" in sampleDF["Test Name"].values:
        clickList = ["check-D10S1248", "check-D12S391", "check-D22S1045",
                     "check-D2S441", "check-D1S1656", "check-DYS391"]
        for each in clickList:
            ActionChains(driver).move_to_element(
                driver.find_element("id", each)).click().perform()

    # Input each allele into the webpage from the dataframe
    for each in master_list:
        marker = each[0]
        row = sampleDF.loc[sampleDF["Markers"] == marker]

        # combine Allele 1-4 from row into a single string, skip if NaN
        alleles = []
        for i in range(1, 5):
            allele = row[f"Allele{i}"].values
            if len(allele) > 0 and not pandas.isnull(allele[0]):
                alleles.append(str(allele[0]))

        allele = ",".join(alleles) if len(alleles) > 0 else ""

        # Input the allele into the webpage
        if len(allele) > 0:
            driver.find_element("id", each[1]).send_keys(allele)

    # Click the search button
    driver.find_element("id", "search").click()

    # Wait for the results to load
    try:
        WebDriverWait(driver, waitTime).until(EC.element_to_be_clickable((By.ID, "export")))
    except Exception as e:
        # print("Timeout for " + sampleName + " from Expasy STR Search")
        pass

    # if no results, return empty list of empty dictionary

    warning_element = driver.find_element("id", "warning")
    display_value = warning_element.value_of_css_property('display')

    if display_value == 'block':
        print("No results for " + sampleName + " from Expasy STR Search")
        data_dict = {
            "Accession": sampleName,
            "Name": "N/A",
            "Nº Markers": "N/A",
            "Score": "0%",
            "Amel": "",
            "CSF1PO": "",
            "D2S1338": "",
            "D3S1358": "",
            "D5S818": "",
            "D7S820": "",
            "D8S1179": "",
            "D13S317": "",
            "D16S539": "",
            "D18S51": "",
            "D19S433": "",
            "D21S11": "",
            "FGA": "",
            "Penta D": "",
            "Penta E": "",
            "TH01": "",
            "TPOX": "",
            "vWA": ""
        }
        # create a pandas dataframe from the data
        tableDF = pandas.DataFrame(data_dict, index=[2])
        # fill index 0 and 1 with None
        tableDF.loc[0] = "None"
        tableDF.loc[1] = "None"
        # Sort the index
        tableDF = tableDF.sort_index()
    else:
        table = driver.find_element("id", "table-results")
        html = table.get_attribute("outerHTML")
        soup = BeautifulSoup(html, "html.parser")
        rows = soup.find_all("tr")
        header_row = rows[0]
        column_names = [th.text for th in header_row.find_all("th")]

        # Store the data in a list of lists
        data = []
        for row in rows:
            cells = row.find_all("td")
            values = [cell.text for cell in cells]
            data.append(values)

        tableDF = pandas.DataFrame(data, columns=column_names)

    if debug:
        print("Expasy Results for " + sampleName + ":")
        print(tableDF)

    # Close the browser
    driver.quit()

    # number of results to return
    number_of_results = 10

    bestMatched = []
    threads = []

    # Create a thread for the first 10, if there are 10, otherwise create a thread for each result

    for i in range(2, number_of_results + 2):
        try:
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
            sampleName, sampleDF, bestMatched[i], "Expasy"))
        threads.append(t)

    # Start the threads
    for t in threads:
        t.start()

    # Wait for the threads to finish
    for t in threads:
        t.join()

    return replacementsDictionaries
