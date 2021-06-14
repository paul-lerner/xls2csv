import argparse
import os
import csv
import sys
from pandas.core.frame import DataFrame
import xlrd
from xlrd.biffh import XLRDError
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True
import pandas as pd

def writecsv (sheet):
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

parser = argparse.ArgumentParser()

#parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\CES FNY Managed Accounts LLC User Detail.xlsx")
parser.add_argument('--source', '-s', help = 'Please, provide the source', type = str, default = "C:\\Users\\Pash\\Downloads\\test_html.xls")
parser.add_argument('--destination', '-d', help = 'Please, provide the destination', type = str, default = "C:\\Users\\Pash\\Downloads\\result.csv")
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

if not (".xls" in str(args.source).lower()):
    sys.stderr.write('Source file does not appear to be an Excel file.')
    sys.stderr.flush
    sys.exit(2)

if not (".csv" in str(args.destination).lower() or ".txt" in str(args.destination).lower()): 
    sys.stderr.write('Output file does not appear to be a CSV/TXT file.')
    sys.stderr.flush
    sys.exit(2)

try: 
#    with xlrd.open_workbook(args.source).raises(XLRDError) as wb:
#        sh = wb.sheet_by_index(0)
#        writecsv(sh)
    data_xls = DataFrame(pd.read_html(args.source))
    data_xls.all
    #data_xls.to_csv(args.destination, encoding='utf-8', index=False)
    #with open(args.destination, 'w', newline="") as f:
    #    col = csv.writer(f)
    #    for x in range(len(data_xls)):
    #        col.writerow(data_xls.__getitem__(x))                
except FileNotFoundError:
    sys.stderr.write('File does not exist')
    sys.stderr.flush
    sys.exit(2)
except PermissionError:
    sys.stderr.write('Could not access the file')
    sys.stderr.flush
    sys.exit(2)
except XLRDError:
    #with pandas.read_html(args.source) as html_wb:
    #    sh = html_wb.sheet_by_index(0)
    #    writecsv(sh)
    sh = pd.read_html (args.source)
    writecsv(sh)
except Exception as e:
    sys.stderr.write('Error: ' + str(e) )
    sys.stderr.flush
    sys.exit(2)
#else:
#finally:
#end: