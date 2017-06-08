# -*- coding: utf-8 -*-
# -*- author: mxiz -*-
import os.path
from gspread import *
from zillow_function import findzillow
from apiclient import discovery
from googleapiclient.errors import HttpError
from oauth2client.service_account import ServiceAccountCredentials
from normal_mode import *
HTD_ADDS = 'https://docs.google.com/spreadsheets/d/1R8QHhkdPOPkgOh8r5Ow0ZoYgkmJop_YISIGAoOXYkbo/edit#gid=0'
MEC_ADDS = 'https://docs.google.com/spreadsheets/d/14KCXK2WlqsDsthv55GVbnpJj3grkvcmGd887aJfse-E/edit#gid=0'
KEY = 'ScrapyCountyWindows-2596dcd5179c.json'

hunterdon = {'name': 'Hunterdon', 'csv': 'hunterdon_items.csv', 'add': HTD_ADDS}
mercer = {'name': 'Mercer', 'csv': 'mercer_items.csv', 'add': MEC_ADDS}
'''
result = service.spreadsheets().values().get(
    spreadsheetId=spreadsheetID, range='C6', valueRenderOption='FORMULA').execute()
print result['values'][0][0]
'''
COUNTY = [hunterdon, mercer]
'''
0             1       2      3        4    5   6    7      8    9
sale_date,sheriff_no,upset,att_ph,case_no,plf, att,address,dfd,schd_data
1      2            3           6          7         10       11        12       13
date, SHERIFF'S #, ADDRESS, Judegment,	NEW UPSET, PLF/DEF, ATTY/FIRM, DOCKET#, Zillow
'''
def manual_mode(num, old_tab_name, worksheet_new_name, startrow):
	county_info = match(num, old_tab_name)
	county = county_info['county']
	filename = county['csv']
	### check .csv file ###
	if os.path.isfile(filename):
		pass
	else:
		print "Error!!!!!"
		print "Not found .csv file.\nPlease run in normal mode first."
		#tkMessageBox.showerror("Error!", "Not found .csv file.\nPlease run in normal mode first.")

	### write manully ###
	try:
		print "--------------------------------------------------"
		print county['name'] + " County starts from NO. " + str(startrow-5) + " item!"
		read_and_write(county_info, worksheet_new_name, startrow)
	except:
		print "--------------------------------------------------"
		print "\t\tNetwork problme. please try again."
		print "--------------------------------------------------"
		#tkMessageBox.showerror("Error!", "Network Problem.\nPlease run it again.")
		quit()

	#tkMessageBox.showinfo("Congrats", "Finished! \nPlease wait until back-up process done.")
	
	### Back Up ###
	back_up(county_info)
	'''
	print "-----------------------------------------------------------"
	print "\t\tBackup Done! Wait for NJlispendens"
	print "-----------------------------------------------------------"
	
	### NJLispenden ###
	njlispendens.njlis_pic(num)
	'''
	print "-----------------------------------------------------------"
	print "\t\tAll Done! Exit anytime."
	print "-----------------------------------------------------------"