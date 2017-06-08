# -*- coding: utf-8 -*-
# -*- author: mxiz -*-
'''
Python 2.7
To use ScrapyCounty crawler:
1. CHECK PIP: If not, download here -> [https://pip.pypa.io/en/stable/installing/] 
2. Scrapy [http://scrapy.org/]
    Install:  $ pip install scrapy
3. Selenium [https://pypi.python.org/pypi/selenium]
    Install: $ pip install selenium
4. PhantomJS
    Install: $ sudo pkg install phantomjs
    [Tip for Windows:
        Change path for PhantomJS first]

To upload to spreadsheet:
1. gspread [Google Spreadsheets Python API https://github.com/burnash/gspread]
    Install: $ pip install gspread
2. Obtain OAuth2 credentials from Google Developers Console
    [http://gspread.readthedocs.io/en/latest/oauth2.html]
3. httplib2
    Install: pip install httplib2
4. Google API client
    [https://developers.google.com/api-client-library/python/start/installation]
5. Before run it, make sure already shared googlesheet with @client
'''
import subprocess
from oauth2client.service_account import ServiceAccountCredentials
import hunterdon_save
import mercer_convert
import datetime
import normal_mode
'''
COUNTY = [morris, essex, bergen, hunterdon, union, mercer, middlesex, monmouth, passaic]
'''
START = 0

def main():
    print "-----------------------------------------"
    print "         Welcome to Scrapy County!"
    print "Choose county:"
    print "              -- Mercer (C)"
    print "              -- Hunterdon (H)"
    print "-----------------------------------------"
    num = raw_input("your choice?")
    name = raw_input('Old Tab Name?')
    if num.upper() == 'C':
        START = datetime.datetime.now()
        mercer(1, name)
    elif num.upper() == 'H':
        START = datetime.datetime.now()
        hunterdon(0, name)
    else:
        print "NOT FOUND! "


    if START is not 0:
        runtime = datetime.datetime.now() - START
        print "-----------------------------------------------------------"
        print "\t\t\tRunning Time is " + str(runtime.seconds) + "s."
        print "-----------------------------------------------------------"
    else:
        print "No......"

def hunterdon(number, name):
    print "Download .pdf ? (Y/N) "
    b = raw_input(": ")
    if b.upper() == 'Y':
        hunterdon_save.hunterdon_save()
    normal_mode.normal_mode(number, name)

def mercer(number, name):
    print "Mercer County is called."
    #os.system("title Mercer County")
    print "Step 1: Download sheriff_foreclosuresales_list.pdf from http://www.mercercounty.org/home/showdocument?id=2246"
    print "Step 2: Open http://www.pdftoexcel.com/"
    print "Step 3: Upload sheriff_foreclosuresales_list.pdf and download .xlsx"
    print "Step 4: Save as sheriff_foreclosuresales_list.xlsx in the same folder"
    print "Please type Y to continue"
    num = raw_input("")
    if num.upper() == 'Y':
        mercer_convert.main()
        normal_mode.normal_mode(number, name)


if __name__ == "__main__":
    main()  