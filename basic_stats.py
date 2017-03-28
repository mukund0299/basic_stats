#Without using a package system, setting everything up in one file
#as separate functions to be used at the user's discretion

import openpyxl
import mpmath
import statistics

#Normal Model
def invNorm(area):
    #Setting initial result to zero
	result = 0
    #The lower bound is set to -3 because it's unlikely to have something
    #higher than that
	i = -3
    #If the result is not between area-0.0001 and area+0.0001
    #Continue to recalculate the area with the new lower bound value
	while (not((result >= (area-0.000001)) and (result <= (area+0.000001)))):
		i += 0.00001
		result = float(0.5*mpmath.erfc(-((i)/mpmath.sqrt(2))))
    #Return the lower bound value
	return i

#Confidence Intervals
def oneSampleZInterval(src, setSheet, confidenceLevel, column):
    #Open the workbook and set the active sheet
    try:
        wb = openpyxl.load_workbook(src)
		sheet = wb.get_sheet_by_name(setSheet)
		index = openpyxl.cell.column_index_from_string(column)
    except FileNotFoundError:
        return 'File not found'
	except KeyError:
		return 'Sheet name not found'

    zCriticalNeg = invNorm((1-(confidenceLevel/100))/2)
    zCritical = -1 * zCriticalNeg
