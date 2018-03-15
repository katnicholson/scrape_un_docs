# Web Scrape: UN Documents
# Scrape (download) all UN statements/speeches by Indian, Pakistani and Bangladeshi Prime Ministersstatements

# (Author: Kathryn Nicholson, Updated: 11/07/2017)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import NoSuchElementException
from optparse import OptionParser
import json, pdb, csv, time, sys, urllib, urllib2, base64, hashlib, hmac, os

# Command line options to specify input output files
parser = OptionParser()

# Input path is the full input path to save .pdfs
parser.add_option("-i", "--input-path", dest="input_path",
                  help="full path to input folder with .csv files with links", metavar='"/mnt/c/Users/username/Documents"')

# Output path is the full output path to save .pdfs
parser.add_option("-o", "--output-path", dest="output_path",
                  help="full path to output folder to save .pdf files", metavar='"/mnt/c/Users/username/Documents"')

# load command line args into parser
(options, args) = parser.parse_args()

# Set folder path for downloads (input path, default if not specified)
if options.input_path:
    input_path = os.path.expanduser(options.input_path)
else:
    input_path = 'C:\\Users\\username\\Documents\\un_scrape'

# Set folder path for downloads (output path, default if not specified)
if options.output_path:
    output_path = os.path.expanduser(options.output_path)
else:
    output_path = 'C:\\Users\\username\\Documents\\UN_pdfs'

# create the output folder if it doesn't exist
if not os.path.exists(output_path): os.mkdir(output_path)

# print initiation messages to user: download folder, start time (stored to var)
print("Download Folder: {}".format(output_path))
process_start_dtime = time.asctime( time.localtime(time.time()) )
print "Process Start: %s \n\n" % process_start_dtime

# set global chromedriver path
chromedriver = "C:\\Users\\username\\Desktop\\chromedriver_win32\\chromedriver.exe"

# list country-specific FAQ pages with links to all UN docs
# Countries: Bangladesh, India, Pakistan
in_link = 'http://ask.un.org/faq/84839'
pk_link = 'http://ask.un.org/faq/75192'
bd_link = 'http://ask.un.org/faq/68198'

# a href=          A/71/PV.22
# pdf lnk format:  http://www.un.org/ga/search/view_doc.asp?symbol=A/71/PV.22
# doc lnk format:  http://daccess-ods.un.org/access.nsf/GetFile?OpenAgent&DS=A/71/PV.22&Lang=E&Type=DOC

# DEFINE FUNCTION: download_pdf
# download embedded Chrome .pdf; arguments = link (alnk = 'A/PV.1582'), download_folder
def download_pdf(alnk, download_folder):
    options = webdriver.ChromeOptions()
    #download_folder = output_path
    profile = {"plugins.plugins_list": [{"enabled": False,
                                         "name": "Chrome PDF Viewer"}],
               "download.default_directory": download_folder,
               "download.extensions_to_open": ""}
    options.add_experimental_option("prefs", profile)
    lnk = 'http://www.un.org/ga/search/view_doc.asp?symbol=' + alnk
    print("Downloading .pdf file from link: {}".format(lnk))
    driver = webdriver.Chrome(executable_path=chromedriver, chrome_options = options)
    driver.get(lnk)
    time.sleep(7)
    print("Status: PDF Download Complete.")
    driver.close()

# DEFINE FUNCTION: download_doc
# convert link to match expression for doc download, download word doc; argument (save as pdf)
def download_doc(alnk, download_folder):
    options = webdriver.ChromeOptions()
    #download_folder = output_path
    profile = {"download.default_directory": download_folder,
               "download.extensions_to_open": ""}
    options.add_experimental_option("prefs", profile)
    lnk = 'http://daccess-ods.un.org/access.nsf/GetFile?OpenAgent&DS=' + alnk + '&Lang=E&Type=DOC'
    print("Downloading .doc file from link: {}".format(lnk))
    driver = webdriver.Chrome(executable_path=chromedriver, chrome_options = options)
    driver.get(lnk)
    time.sleep(5)
    # test for error in document link
    print "Testing for document link error..."
    try:
        element = driver.find_element_by_xpath('//h4')
        text = element.get_attribute('innerHTML')
        if "Error" in text:
            print "ERROR: Unable to download Word document."
            driver.close()
            # (write failed link to dictionary) download .pdf
            download_pdf(alnk, download_folder)
        else:
            pass
    except:
        print("Status: WORD DOC Download Complete.")
        driver.close()

###############
## MAIN CODE ##
###############

# open each country .csv file and click on links stored one-per-row
for root, dirs, files in os.walk(input_path):
    for f in files:
        if f[-3:].lower() == "csv":

            # display country
            print "Processing: %s \n" % f[0:-4]

            # initiate counter
            downloads = 0

            # create country folders under output folder for downloads
            output_folder = os.path.join(output_path, f[0:-4])
            if not os.path.exists(output_folder): os.mkdir(output_folder)

            # open csv file and read each row
            f_open = open(os.path.join(input_path, f))
            file_read = csv.reader(f_open)

            # Go to alnk in each row: download_doc tries doc first, then pdf
            for row in file_read:

                # download PDF/DOC
                download_doc(row[0], output_folder)
                
                # print the newest file in the directory to user
                # os.system('ls %s -t * | head -1' % output_folder)
                # os.system('dir %s /d /o:-d /t:w /a:-d' % output_folder)
                os.system('dir %s /d /o:-d /t:w /a:-d | find ".pdf" | findstr /n ^^ | findstr "^1"' % output_folder)
                print "\n"
                
                downloads += 1

            # print # of files downloaded
            print "Total downloads: %d \n" % downloads

            # close .csv after reading last line
            f_open.close()

# print time logs
print "Finished! Processed %d total pdf requests \n" % downloads
print "Process Start: %s" % process_start_dtime
print "Process End: %s" % time.asctime( time.localtime(time.time()) )

# exit python
quit()

#########################
#### EXTRA FUNCTIONS ####
#########################

# Test for errors to troubleshoot between .docx, .pdf link, or extractable HTML text

# Example, if .pdf link is wrong, error on page: "There is no document"

# ex good: driver.get('http://daccess-ods.un.org/access.nsf/GetFile?OpenAgent&DS=A/71/PV.22&Lang=E&Type=DOC')
# ex fail: driver.get('https://daccess-ods.un.org/access.nsf/GetFile?OpenAgent&DS=A/PV.1582&Lang=E&Type=DOC')

# DEFINE FUNCTION: test_doc_error
# test for failed document download links (old UN years), Error 91: Object variable not set
def test_doc_error(driver):
    element = driver.find_element_by_xpath('//h4')
    text = element.get_attribute('innerHTML')
    if "Error" in text:
        print "ERROR: Unable to download Word document"
        # write failed link to dictionary and download .pdf
        download_pdf(lnk)
    else:
        pass
		
# DEFINE FUNCTION: convert_pdf_link
# format text matching a='' on page to navigate to pdf view
def convert_pdf_link(lnk_text):
    lnk = 'http://www.un.org/ga/search/view_doc.asp?symbol=' + lnk_text
    return lnk

# DEFINE FUNCTION: convert_doc_link
# format text matching a='' on page to navigate to doc download
def convert_doc_link(lnk_text):
    lnk = 'http://daccess-ods.un.org/access.nsf/GetFile?OpenAgent&DS=' + lnk_text + '&Lang=E&Type=DOC'
    return lnk