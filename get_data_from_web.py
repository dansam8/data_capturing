'''
outputs data in spreadsheet name: output.xlsx format:

subburb, street, erf, street No, ownersName, Extent, Erfusage
'''

import sys, os, time, cv2, xlrd
from pykeyboard import PyKeyboard
from pymouse import PyMouse
from PIL import ImageGrab
from openpyxl import Workbook
from openpyxl import load_workbook
from difflib import SequenceMatcher


pathToRoadList = '/input.txt'
pathToOutput = '/output.xlsx'
pathToDownloadXls = '/Users/DWDSamuelson/Downloads/StreetOwner.xls'
suburbSimilarityThreshhold = 0.8


class Sudo_scraper:

	def is_pix_yellow(x,y):
		img = np.array(ImageGrab.grab().convert('RGB'))
		d1,d2,d3 = img[y,x] == [254, 255, 206]
		return d1 and d2 and d3

	def search(road):
		time.sleep(1)
		m.click(750,400)
		time.sleep(1)
		for i in range(40):
			k.tap_key('delete')
		k.type_string(road)
		time.sleep(1)
		k.tap_key('return')
		time.sleep(8)

	def check_if_result_exists_at(resultnumber = 0):
		y = 270+(resultnumber*37)
		x = 500
		m.move(x,y)
		time.sleep(0.5)
		return(yellow(x,y))

	def download_result_at(resultnumber=0):
		m.click(500,(270+resultnumber*37))
		m.move(600,200)
		time.sleep(6)	
		m.click(1025,200)
		time.sleep(0.2)
		m.click(1025,200)
	 	time.sleep(15)
		m.click(1072,197)
		m.click(1072,197)
		time.sleep(1)
		m.click(1072,198)
		time.sleep(1.5)
		m.click(1072,198)
	def exit_results_window():
		time.sleep(2)
		m.click(920,198)

def openxls(path,sheetindex=0):
	while True:
		if os.path.exists(path):
			try:
				f = xlrd.open_workbook(path)
				sheet = f.sheet_by_index(sheetindex)
				return sheet
			except Exception as e:
				print e
				raw_input("Error")
		else:
			raw_input("file is not at path "+path+" Enter to try again \nerror from function openxls")

def opentxt(path,mode='r+a'):
	while True:
		if os.path.exists(path):
			try:
				f = open(path,mode)
			except Exception as e:
				print e
				raw_input("Error opening txt file "+path+" \n error from function opentxt")


def String_difference_value(one,two): # between 0-1
	return SequenceMatcher(None, one.lower(), two.lower()).ratio()

def split_string_by_delimiter(string, delimiter=','):
	one = inp[:inp.index(delimiter)].strip()
	two = inp[inp.index(delimiter)+1:].strip()
	return[one,two]

def remove_file(path):
	if os.path.exists(path):

		try:
			os.remove(path)
			return True
		except Exception as e:
			print e
			raw_input("file Error")
	else:
		raw_input("file not at path "+path+" \nerror from function remove_file")

def get_suburb_from_file(path):
	sheet = openxls(path)
	st = sheet.cell(1,0).value
	return st[st.index(',')+2:-1]

def check_input_file_structure():
	pass


def test(pathToDownloadXls,pathToOutput,pathToRoadList):
	if os.path.exists(pathToDownloadXls):
		print "download folder alrealy contains file for downloading "+pathToDownloadXls
		exit()
	if os.path.exists(pathToOutput):
		print "output file already exists "+pathToOutput
		exit()
	if not os.path.exists(pathToRoadList):
		print "can't find input roadlists file "+pathToRoadList
		exit()

	check_input_file_structure()

def find_sheet_from_resutls(road,suburb,scraper,suburbSimilarityThreshhold,pathToDownloadXls):#returns true for false
	count = 0
	while True:
		scraper.search(road)
		if scraper.check_if_result_exists_at(count):
			scraper.download_result_at(count)
			actualSuburb = get_suburb_from_file(pathToDownloadXls)
			if suburbSimilarityThreshhold < String_difference_value(actualSuburb,suburb):
				return True
		else:
			break
	
	return False
def add_data_to_output_sheet():
	pass

def get_number_of_rows_in_output_file():
	pass

def get_data(pathToDownloadXls,pathToOutput,pathToRoadList,suburbSimilarityThreshhold):

	test(pathToDownloadXls,pathToOutput,pathToRoadList)

	print "test complete"
	for i in range(5):
		print "starting scraping in "+5-i
		time.sleep(1)

	sudoScraper = Sudo_scraper()

	inputRoadsFile = opentxt(pathToRoadList)

	for line in inputRoadsFile.readlines():
		if '##' in line:
			print "complete"
			exit()
		else:
			road,suburb = split_string_by_delimiter(line)
			if find_sheet_from_resutls(road,suburb,sudoScraper,suburbSimilarityThreshhold,pathToDownloadXls):
				print "True "+line
				f.write("True "+line)
				add_data_to_output_sheet()
			else:
				print "False "+line+" at line "+get_number_of_rows_in_output_file() # for posibly adding data later
				f.write("False "+line+"at line "+get_number_of_rows_in_output_file())	

























