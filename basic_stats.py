#Without using a package system, setting everything up in one file
#as separate functions to be used at the user's discretion

import openpyxl
import sympy

#Normal Model
def invNorm(x):


#Confidence Intervals
def oneSampleZInterval(src, setSheet, confidenceLevel, column):
    try:
        wb = openpyxl.load_workbook(src)
    except FileNotFoundError:
        return 'File not found'
    zCritical = invNorm((1-(confidenceLevel/100))/2)
    
