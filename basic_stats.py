#Without using a package system, setting everything up in one file
#as separate functions to be used at the user's discretion

import openpyxl

#Confidence Intervals
def oneSampleZInterval(src, sheet, confidenceLevel, column):
    try:
        wb = openpyxl.load_workbook(src)
    except FileNotFoundError:
        return 'File not found'
    
