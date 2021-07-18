import argparse
import os
import csv
import sys
import xlrd
import pandas as pd
from pandas.core.frame import DataFrame
from xlrd.biffh import DATEFORMAT, XLRDError
from xlrd.formula import dump_formula
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

#find maximum number of rows in a file
def findmaxcolumncount(DataFrame):
	maxlength=0 
	#print(DataFrame)
	for data in DataFrame:
		for value in data.values:
			if len(value) > maxlength:
				maxlength = len(value)
	return maxlength

#Read HTML file disguised as an Excel file
def readhtmlfile(filepath):
	try:
		data_xls = pd.read_html(filepath,parse_dates=True,na_values='')
		return data_xls
	except FileNotFoundError:
		sys.stderr.write('File does not exist')
		sys.stderr.flush
		sys.exit(2)
	except IOError:
		sys.stderr.write('File error')
		sys.stderr.flush
		sys.exit(2)

#Read Tab delimited file disguised as an Excel file
def readtextfile(filepath, delimeter=','):
	try:
#		data_xls = pd.read_table(filepath, error_bad_lines=False)
		with open(filepath,'r',newline='') as tab_temp:                                                                                          
			datafile = tab_temp.readlines()
			temparray1 = []
			maxlength = 0
			for line in datafile:
				temparray = line.split(delimeter)
				temparray = [item.replace('\n', '') for item in temparray]
				if len(temparray1) > maxlength:
					maxlength = len(temparray)
				list.insert(temparray1,len(temparray1),temparray)
			temparray2 = []
			for temp in temparray1:
				while (len(temp) <= maxlength):
							temp.append('')
							temp = [item.replace('\r', '') for item in temp]
				list.insert(temparray2,len(temparray2),temp)
			data_xls = temparray2
			return data_xls
	except FileNotFoundError:
		sys.stderr.write('File does not exist')
		sys.stderr.flush
		sys.exit(2)
	except IOError:
		sys.stderr.write('File error')
		sys.stderr.flush
		sys.exit(2)

def writetabfile(sheet,filepath):
	try:
		with open(filepath, 'w',newline='') as result_file:
			result_writer = csv.writer(result_file, dialect=csv.excel, quoting=csv.QUOTE_MINIMAL)
			result_writer.writerows(sheet)
	except FileNotFoundError:
		sys.stderr.write('File does not exist')
		sys.stderr.flush
		sys.exit(2)
	except IOError:
		sys.stderr.write('File error')
		sys.stderr.flush
		sys.exit(2)


#Check if source is a Tab delimited file disguised as an Excel file
def itsatabfile(filepath):
	try:
		with open(filepath) as temp_f:
			datafile = temp_f.readlines()
		for line in datafile:
			if line.count('\t') > 1:
				return True # The string is found
		return False
	except FileNotFoundError:
		sys.stderr.write('File does not exist')
		sys.stderr.flush
		sys.exit(2)
	except IOError:
		sys.stderr.write('File error')
		sys.stderr.flush
		sys.exit(2)

def writecsvsheet (sheet):
	try:
		with open(args.destination, 'w', newline="") as f:
			col = csv.writer(f)
			for row in range(sheet.nrows):
				col.writerow(sheet.row_values(row))
	except FileNotFoundError:
		sys.stderr.write('File does not exist')
		sys.stderr.flush
		sys.exit(2)
	except IOError:
		sys.stderr.write('File error')
		sys.stderr.flush
		sys.exit(2)

def writecsv(data_xls,filepath,maxlength=-2):
	try:
		with open(filepath, 'w',newline='') as result_file:
			result_writer = csv.writer(result_file, dialect=csv.excel, quoting=csv.QUOTE_MINIMAL)
			temparray1 = []
			for data in data_xls:
				for value in data.values:
						temparray = list(value)
						i=0
						for element in temparray:
							if str(element).__eq__("nan"):
								temparray[i] = ""
							i=i+1
						if maxlength > 0:
							while (len(temparray) <= maxlength):
								temparray.append(" ")
						list.insert(temparray1,len(temparray1),temparray)
			#    print(temparray1)
			result_writer.writerows(temparray1)
	except FileNotFoundError:
		sys.stderr.write('File does not exist')
		sys.stderr.flush
		sys.exit(2)
	except IOError:
		sys.stderr.write('File error')
		sys.stderr.flush
		sys.exit(2)

parser = argparse.ArgumentParser()

#parser.add_argument('--destination', '-d', help = 'Please, provide the destination', type = str, default = "C:\\Users\\Pash\\Downloads\\result.csv")
#parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\CES FNY Managed Accounts LLC User Detail.xlsx")
#parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\result1.csv")
#parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\test_html.xls")
#parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\test_tab.xls")
#parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\test_nan.xls")
#parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\test_date.xlsx")

parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str)
parser.add_argument('--destination', '-d', help = 'Please, provide the destination', type = str)
parser.add_argument('--remove', '-r', help = 'Please, if source to be removed', type = bool, default=False)
parser.add_argument('--overwrite', '-o', help = 'Please, if destination to be overwritten', type = bool, default=False)

args = parser.parse_args()

print(parser.format_help())
#usage: convertXLS2CSV [-h] [--source=full path to source] [--destination=full path to destination ]
#Source file is expected to be a properly formatted Excel file with .xls or .xlsx extension.
#Output file is expected to have .csv or .txt extension.
#
#optional arguments:
#   -s, --source        full path to source file or folder
#   -d, --destination   full path to destination file or folder
#   -h, --help          show this help message
#   -o, --overwrite     overwrite destination file if exist

if not (".xls" in str(args.source).lower() or ".csv" in str(args.source).lower()):
	sys.stderr.write('Source file does not appear to be an Excel file.')
	sys.stderr.flush
	sys.exit(2)

if not (".csv" in str(args.destination).lower() or ".txt" in str(args.destination).lower()): 
	sys.stderr.write('Output file does not appear to be a CSV/TXT file.')
	sys.stderr.flush
	sys.exit(2)

if (".csv" in str(args.source).lower()):
	writetabfile(readtextfile(args.source,','),args.destination)
else:
	try:
		Workbook = pd.read_excel(args.source)
		Workbook.to_csv(args.destination)
#		WorkBook = xlrd.open_workbook(args.source)
#		Sheet = WorkBook.sheet_by_index(0)
#		with open(args.destination, 'w', encoding='utf8',newline='') as your_csv_file:
#			wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

#			for rownum in range(Sheet.nrows):
#				wr.writerow(Sheet.row_values(rownum))


	except FileNotFoundError:
		sys.stderr.write('File does not exist')
		sys.stderr.flush
		sys.exit(2)
	except PermissionError:
		sys.stderr.write('Could not access the file')
		sys.stderr.flush
		sys.exit(2)
	except XLRDError:
		if itsatabfile(args.source):
			data_xls = readtextfile(args.source, '\t')
			if len(data_xls) > 0:
				writetabfile(data_xls,args.destination)
		else:
			data_xls = readhtmlfile(args.source)
			if len(data_xls) > 0:
				writecsv(data_xls,args.destination,findmaxcolumncount(data_xls))    
	except Exception as e:
		sys.stderr.write('Error: ' + str(e) )
		sys.stderr.flush
		sys.exit(2)
	#else:
	#finally: