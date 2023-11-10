# For the Clima Website Contact information
import platform
email_key = "info@laragen.com"

# Selenium Parameters
# Set the wait time for the browser to load the page in seconds
# This is used to prevent the script from running too fast and causing errors
waitTime = 20

# Set the number of threads to use for processing the samples
# Too many threads will cause the Clima or Expasy server to not respond quickly enough
MaxThreads = 2

# Allows Selenium to run in the background without opening a browser window
Headless = True

# File Paths

# For Development and Debugging Purposes
debug = False

# If the program is Not running on ETNA then set Headless to True and debug to False automatically
if platform.node() != "ETNA":
    Headless = True
    debug = False
