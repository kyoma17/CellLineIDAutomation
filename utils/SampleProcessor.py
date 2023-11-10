import pandas as pd
from params import MaxThreads
from utils.ExpasyScraper import ExpasySTRSearch
from utils.ClimaScraper import ClimaSTRSearch
import warnings
import threading
import queue
import concurrent.futures
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)


def processSamples(df, max_threads=MaxThreads):
    '''
    Takes in a pandas dataframe and processes each sample
    Args:
        df: The dataframe to process
        max_threads: The maximum number of threads to use
    Returns:
        result_collection: The collection of results
        sample_order: The order of the samples
    '''

    df = remove_trailing_zeros(df)
    remove_OL(df)

    sampleList = df["Sample Name"].unique()

    grouped_df = df.groupby("Sample Name", sort=False)

    print("Processing " + str(len(sampleList)) + " samples..." + '\n')

    sample_counter = 0
    sample_order = []

    result_collection = []

    # Create a semaphore
    semaphore = threading.Semaphore(max_threads)

    # Selenium Script for Multi Threaded Processing
    def process_sample(sampleName, sampleDF, sample_counter, results_queue, semaphore):
        # Acquire a semaphore
        semaphore.acquire()

        testName = sampleDF["Test Name"].values

        if "GenePrint_24_POP7_Panels_v1.0" in testName:
            print("Processing GP24 " + sampleName + "...")
        else:
            print("Processing GP10 " + sampleName + "...")

        # expasy_results = ExpasySTRSearch(sampleName, sampleDF, sample_counter)
        # clima_results = ClimaSTRSearch(sampleName, sampleDF, sample_counter)

        # Create a ThreadPoolExecutor with 2 threads
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:

            # Submit the tasks to the executor
            expasy_future = executor.submit(ExpasySTRSearch, sampleName, sampleDF, sample_counter)
            clima_future = executor.submit(ClimaSTRSearch, sampleName, sampleDF, sample_counter)

            # Get the results from the futures
            expasy_results = expasy_future.result()
            clima_results = clima_future.result()

        # Combine the results from ClimaSTR and ExpasySTR
        results = clima_results + expasy_results

        # Add the results to the results queue
        results_queue.put([results, sampleName])

        # Release the semaphore
        semaphore.release()

    def process_grouped_df(grouped_df):
        sample_counter = 0
        results_queue = queue.Queue()
        threads = []

        for each in grouped_df:
            sample_counter += 1
            sampleName = each[0]
            sample_order.append(sampleName)
            sampleDF = each[1]

            thread = threading.Thread(target=process_sample, args=(sampleName, sampleDF, sample_counter, results_queue, semaphore))
            thread.start()
            threads.append(thread)

        # Wait for all threads to finish
        for thread in threads:
            thread.join()

        # Retrieve the results from the queue in the original order
        while not results_queue.empty():
            result_collection.append(results_queue.get())

    process_grouped_df(grouped_df)

    return result_collection, sample_order


#####################################################################################################
# Helper Functions
#####################################################################################################


def check_allele_empty(df):
    '''
    Checks if all allele columns are empty
    '''
    # Define columns to process
    allele_cols = ['Allele1', 'Allele2', 'Allele3', 'Allele4']

    # check if all allele columns are empty
    if df[allele_cols].isnull().all().all():
        return True
    else:
        return False


def remove_trailing_zeros(df):
    '''
    Removes trailing zeros from the allele columns
    '''
    # Define columns to process
    allele_cols = ['Allele1', 'Allele2', 'Allele3', 'Allele4']

    # Process each column
    for col in allele_cols:
        # Convert numbers with trailing .0 to integer, but keep NaN as NaN
        df[col] = df[col].apply(lambda x: str(x).replace('.0', '') if pd.notna(x) else x)

    return df


def remove_OL(df):
    '''
    Removes "OL" from the allele columns and replaces it with an empty string
    '''

    allele_cols = ['Allele1', 'Allele2', 'Allele3', 'Allele4']

    # Process each column
    for col in allele_cols:
        df[col] = df[col].apply(lambda x: str(x).replace('OL', '') if pd.notna(x) else x)


############################################################################################################
    # Selenium Script for Single Threaded Processing
    # for each in grouped_df:
    #     sample_counter += 1

    #     # sample_counter = "#"
    #     # if Test Name is geneprint24

    #     sampleName = each[0]
    #     sample_order.append(sampleName)
    #     sampleDF = each[1]
    #     testName = each[1]["Test Name"].values

    #     if "GenePrint_24_POP7_Panels_v1.0" in testName:
    #         print("Processing GP24 " + sampleName + "...")
    #     else:
    #         print("Processing GP10 " + sampleName + "...")

    #     expasy_results = ExpasySTRSearch(sampleName, sampleDF, sample_counter)
    #     clima_results = ClimaSTRSearch(sampleName, sampleDF, sample_counter)

    #     # Combine the results from ClimaSTR and ExpasySTR
    #     results = clima_results + expasy_results

    #     # Add the results to the result collection for bulk selection
    #     result_collection.append([results, sampleName])
        # print('\n')


# def processSamplesRetired(df):
#     sampleList = df["Sample Name"].unique()

#     grouped_df = df.groupby("Sample Name", sort=False)

#     sample_counter = 0
#     sample_order = []

#     result_collection = []
#     # Selenium Script for Multi Threaded Processing

#     def process_sample(sampleName, sampleDF, sample_counter, results_queue):
#         testName = sampleDF["Test Name"].values

#         if "GenePrint_24_POP7_Panels_v1.0" in testName:
#             print("Processing GP24 " + sampleName + "...")
#         else:
#             print("Processing GP10 " + sampleName + "...")

#         expasy_results = ExpasySTRSearch(sampleName, sampleDF, sample_counter)
#         clima_results = ClimaSTRSearch(sampleName, sampleDF, sample_counter)

#         # Combine the results from ClimaSTR and ExpasySTR
#         results = clima_results + expasy_results

#         # Add the results to the results queue
#         results_queue.put([results, sampleName])

#     def process_grouped_df(grouped_df):
#         sample_counter = 0
#         results_queue = queue.Queue()
#         threads = []

#         for each in grouped_df:
#             sample_counter += 1
#             sampleName = each[0]
#             sample_order.append(sampleName)
#             sampleDF = each[1]

#             thread = threading.Thread(target=process_sample, args=(sampleName, sampleDF, sample_counter, results_queue))
#             thread.start()
#             threads.append(thread)

#         # Wait for all threads to finish
#         for thread in threads:
#             thread.join()

#         # Retrieve the results from the queue in the original order
#         while not results_queue.empty():
#             result_collection.append(results_queue.get())

#     process_grouped_df(grouped_df)

#     return result_collection, sample_order
