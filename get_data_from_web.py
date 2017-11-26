'''
outputs data in spreadsheet name: output.xlsx structure:

subburb, street, erf, street No, ownersName, Extent, Erfusage
'''

import sys, os, time, cv2, xlrd, openpyxl
from pykeyboard import PyKeyboard
from pymouse import PyMouse
from PIL import ImageGrab
from difflib import SequenceMatcher
import numpy as np

pathToRoadList = os.getcwd()+'/input.txt'
pathToOutput = os.getcwd()+'/output.xlsx'
pathToDownloadXls = '/Users/DWDSamuelson/Downloads/StreetOwner.xls'
suburbSimilarityThreshhold = 0.8
output_arr_as_temp_storage = []


class pseudo_scraper:

	def __init__(self):
		self.m = PyMouse()
		self.k = PyKeyboard()

	def is_pix_yellow(self,x,y):
		img = np.array(ImageGrab.grab().convert('RGB'))
		d1,d2,d3 = img[y,x] == [254, 255, 206]
		return d1 and d2 and d3

	def search(self,road):
		time.sleep(1)
		self.m.click(750,400)
		time.sleep(1)
		for i in range(40):
			self.k.tap_key('delete')
		self.k.type_string(road)
		time.sleep(1)

	def check_if_result_exists_at(self,resultnumber = 0):
		y = 270+(resultnumber*37)
		x = 500
		self.m.move(x,y)
		time.sleep(0.5)
		return(self.is_pix_yellow(x,y))

	def download_result_at(self,resultnumber=0):
		self.m.click(500,(270+resultnumber*37))
		self.m.move(600,200)
		time.sleep(6)	
		self.m.click(1025,200)
		time.sleep(0.2)
		self.m.click(1025,200)
	 	time.sleep(15)
		self.m.click(1072,197)
		self.m.click(1072,197)
		time.sleep(1)
		self.m.click(1072,198)
		time.sleep(1.5)
		self.m.click(1072,198)

	def exit_results_window(self):
		time.sleep(2)
		m.click(920,198)


def save_array_to_xlsx(path):
	global output_arr_as_temp_storage

	while True:
		try:
			wb = openpyxl.Workbook(path)
			wb.create_sheet('sheet')
			sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])
			break
		except Exception as e:
			print e
			raw_input("Error save_array_to_xlsx")

	for line in range(len(output_arr_as_temp_storage)):
		temp = []
		for cell in range(len(output_arr_as_temp_storage[line])):
			temp.append(output_arr_as_temp_storage[line][cell])
		sheet.append(temp)

	wb.save(pathToOutput)


def openxls(path,sheetindex=0):
	while True:
		if os.path.exists(path):
			try:
				f = xlrd.open_workbook(path)
				sheet = f.sheet_by_index(sheetindex)
				return sheet
			except Exception as e:
				print e
				raw_input("Error openxls")
		else:
			raw_input("file is not at path "+path+" Enter to try again \nerror from function openxls")


def opentxt(path,mode='r+a'):
	while True:
		if os.path.exists(path):
			try:
				f = open(path,mode)
				return f
			except Exception as e:
				print e
				raw_input("Error opening txt file "+path+" \n error from function opentxt")
		else:
			raw_input("can't find input data file")


def String_difference_value(one,two): # between 0-1
	return SequenceMatcher(None, one.lower(), two.lower()).ratio()


def split_string_by_delimiter(string, delimiter=','):
	one = string[:string.index(delimiter)].strip()
	two = string[string.index(delimiter)+1:].strip()
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


def find_sheet_from_resutls(road, suburb, suburbSimilarityThreshhold, pathToDownloadXls):#returns true for false
	count = 0
	scraper = pseudo_scraper()
	
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


def add_data_to_output_arr(pathToDownloadXls):
	global output_arr_as_temp_storage

	downloaded_sheet = openxls(pathToDownloadXls)
	for i in range(6,downloaded_sheet.nrows):
		output_arr_as_temp_storage.append([])
		for j in range(downloaded_sheet.ncols):
			output_arr_as_temp_storage[len(output_arr_as_temp_storage)-1].append(str(downloaded_sheet.cell(i,j).value))
	

def get_data(pathToDownloadXls, pathToOutput, pathToRoadList, suburbSimilarityThreshhold):

	test(pathToDownloadXls,pathToOutput,pathToRoadList)

	print "test complete"
	for i in range(5):
		print "starting scraping in "+str(5-i)
		time.sleep(1)
	
	inputRoadsFile = opentxt(pathToRoadList)
	
	for line in inputRoadsFile.readlines():
		if '##' in line:
			print "complete"
			exit()
		else:
			road,suburb = split_string_by_delimiter(line)
			if find_sheet_from_resutls(road, suburb, suburbSimilarityThreshhold, pathToDownloadXls):
				print "True "+line
				inputRoadsFile.write("True "+line)
				add_data_to_output_arr(pathToDownloadXls)
			else:
				print "False "+line
				inputRoadsFile.write("False "+line)

	save_array_to_xlsx(pathToOutput)



get_data(pathToDownloadXls,pathToOutput,pathToRoadList,suburbSimilarityThreshhold)






















