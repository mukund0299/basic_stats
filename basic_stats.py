#Without using a package system, setting everything up in one file as separate functions to be used at the user's discretion

#TODO: One Sample Z Test
#TODO: Two Sample Z Test
#TODO: One Sample T Test
#TODO: Two Sample T Test
#TODO: Two Variable statistics
#TODO: Linear Regressions
#TODO: Plotting (Scatterplots, linear Regressions)
#TODO: Test all functions

#IDEA: Create a GUI for the user rather than CLI

import openpyxl
import mpmath
import statistics
import matplotlib

#Normal Model
def invNorm(area):
    #Setting initial result to zero
	result = 0
    #The lower bound is set to -3 because it's unlikely to have something higher than that
	i = -3
    #If the result is not between area-0.0001 and area+0.0001, continue to recalculate the area with the new upper bound value
	while (not((result >= (area-0.000001)) and (result <= (area+0.000001)))):
		i += 0.00001
		result = float(0.5*mpmath.erfc(-((i)/mpmath.sqrt(2))))
    #Return the upper bound value
	return i

#t distribution
def invt(area, df):
	#Setting the initial result to zero
	result = 0
	#Starting at -4 because t values tend to go out further
	i = -4
	#Same as invNorm
	while (not((result >= (area-0.000001)) and (result <= (area+0.000001)))):
		i += 0.00001
		if (i <= 0):
			x2 = (df)/((i**2)+df)
			result = 0.5*mpmath.betainc((df/2),0.5,0,x2, regularized = True)
		else:
			x2 = (i**2)/((i**2)+df)
			result = 0.5*(mpmath.betainc(0.5,(df/2),0,x2, regularized = True)+1)
	return i

def sumData(data, raiseTo):
	dataTemp = []
	for value in data:
		dataTemp.append(value)
	sumOfValues = 0
	for i in range(len(dataTemp)):
		dataTemp[i] = dataTemp[i]**raiseTo
	for value in dataTemp:
		sumOfValues += value
	return sumOfValues

#Confidence Intervals for proportions
def oneSampleZInterval(successes, n, confidenceLevel):
	#Calculating p from the sample
	pHat = successes/n
	#Getting Critical value
	zCritical = -1*invNorm((1-(confidenceLevel/100))/2)
	#Getting Standard error
	stDev = float(mpmath.sqrt(((pHat)*(1-pHat))/(n)))
	marError = stDev*zCritical
	CLower = pHat - marError
	CUpper = pHat + marError
	print('p-Hat = ' + str(pHat))
	print('ME = ' + str(marError))
	print('CLower = ' + str(CLower))
	print('CUpper = ' + str(CUpper))

def twoSampleZInterval(successes1, n1, successes2, n2, confidenceLevel):
	#Calculating p from the samples
	pHat1 = successes1/n1
	pHat2 = successes2/n2
	#Getting the difference in p
	pDiff = pHat1-pHat2
	#Getting critical value
	zCritical = -1*invNorm((1-(confidenceLevel/100))/2)
	#Getting Standard error
	stDev = mpmath.sqrt((((pHat1)*(1-pHat1))/n1) + (((pHat2)*(1-pHat2))/n2))
	marError = zCritical*stDev
	CLower = pDiff - marError
	CUpper = pDiff + marError
	print('p-Hat1 = ' + str(pHat1))
	print('pHat2 = ' + str(pHat2))
	print('pDiff = ' + str(pDiff))
	print('ME  = ' + str(pDiff))
	print('Clower = ' + str(CLower))
	print('CUpper = ' + str(CUpper))

#Confidence Intervals for means; reading from an excel file
def oneSampleTInterval(src,setSheet,confidenceLevel,column):
	try:
		#Opening the workbook and setting the sheet
		wb = openpyxl.load_workbook(src)
		sheet = wb.get_sheet_by_name(setSheet)
		#Getting a list of the cells in order to calculate length
		dataList = list(sheet.columns)[column]
		#Populating a list with data from the cells
		data = []
		for cellObj in dataList:
			data.append(float(cellObj.value))
		#Getting mean,Standard Deviation,n and df
		mean = statistics.mean(data)
		sampleStDev = statistics.stdev(data)
		n = len(dataList)
		df = n-1
		stDev = (sampleStDev)/(mpmath.sqrt(n))
		#Getting critical value
		tCritical = -1*invt(((1-(confidenceLevel/100))/2),df)
		stError = stDev*tCritical
		CLower = mean - stError
		CUpper = mean + stError
		print('Mean = ' + str(mean))
		print('Standard Deviation = ' + str(stDev))
		print('Critical Value = ' + str(tCritical))
		print('Standard Error = ' + str(stError))
		print('Lower limit = ' + str(CLower))
		print('Upper limit = ' + str(CUpper))
	except FileNotFoundError:
		print('File not Found')
	except KeyError:
		print('Sheet does not exist')
	except TypeError:
		print('File contains letters or empty cells')


def twoSampleTInterval(src, setSheet, confidenceLevel, column1, column2):
	try:
		#Opening the workbook and setting the sheet
		wb = openpyxl.load_workbook(src)
		sheet = wb.get_sheet_by_name(setSheet)
		#Getting a list of the cells in order to calculate length
		dataList1 = list(sheet.columns)[column1]
		dataList2 = list(sheet.columns)[column2]
		#Populating a list with data from the cells
		data1, data2 = [], []
		for cellObj in dataList1:
			data1.append(float(cellObj.value))
		for cellObj in dataList2:
			data2.append(float(cellObj.value))
		#Getting mean,Standard Deviation,n and df for both samples
		mean1 = statistics.mean(data1)
		sampleStDev1 = statistics.stdev(data1)
		n1 = len(dataList1)
		df1 = n1-1
		mean2 = statistics.mean(data2)
		sampleStDev2 = statistics.stdev(data2)
		n2 = len(dataList2)
		df2 = n2-1
		meanDiff = mean1-mean2
		#BUG: The df calculation is wrong
		df = (((sampleStDev1**2)/n1)+((sampleStDev2**2)/n2)**2)/((((sampleStDev1**2)/n1)/df1)+(((sampleStDev2**2)/n2)/df2))
		stDev = float(mpmath.sqrt(((sampleStDev1**2)/n1)+((sampleStDev2**2)/n2)))
		#Getting critical value
		tCritical = -1*invt(((1-(confidenceLevel/100))/2),df)
		stError = stDev*tCritical
		CLower = meanDiff - stError
		CUpper = meanDiff + stError
		print('Clower = ' + str(CLower))
		print('CUpper = ' + str(CUpper))
		print('Difference in mean = ' + str(meanDiff))
		print('SE = ' + str(stError))
		print('df = ' + str(df))
		print('mean1 = ' + str(mean1))
		print('mean2 = ' + str(mean2))
		print('stDev1 = ' + str(sampleStDev1))
		print('stDev2 = ' + str(sampleStDev2))
		print('n1 = ' + str(n1))
		print('n2 = ' + str(n2))
	except FileNotFoundError:
		print('File not Found')
	except KeyError:
		print('Sheet does not exist')
	except TypeError:
		print('File contains letters or empty cells')

def oneVarStats(src, setSheet, column):
	try:
		#Opening the workbook and setting the sheet
		wb = openpyxl.load_workbook(src)
		sheet = wb.get_sheet_by_name(setSheet)
		#Getting a list of the cells in order to calculate length
		dataList = list(sheet.columns)[column]
		#Populating a list with data from the cells
		data = []
		for cellObj in dataList:
			data.append(float(cellObj.value))
		mean = statistics.mean(data)
		sampleStDev = statistics.stdev(data)
		populationStDev = statistics.pstdev(data)
		n = len(data)
		sumOfValues = sumData(data,1)
		sumOfValuesSquared = sumData(data,2)
		#Five Number summary
		data.sort()
		minVal = data[0]
		q1Val, medianVal, q3Val = 0, 0, 0
		maxVal = data[len(data)-1]
		#If the data has odd number of values
		if (len(data) % 2 != 0):
			middleIndex = (len(data)-1)//2
			medianVal = data[middleIndex]
			lowerData = data[0:middleIndex]
			q1Val = statistics.median(lowerData)
			upperData = data[middleIndex+1:len(data)]
			q3Val = statistics.median(upperData)
		else:
			medianVal = statistics.median(data)
			lowerData = data[0:len(data)//2]
			q1Val = statistics.median(lowerData)
			upperData = data[len(data)//2:len(data)]
			q3Val = statistics.median(upperData)
		#Sum of squared deviations
		ssx = 0
		for value in data:
			ssx += (value-mean)**2
		print('Mean = ' + str(mean))
		print('Sum of x = ' + str(sumOfValues))
		print('Sum of squared x = ' + str(sumOfValuesSquared))
		print('Sample standard deviation = ' + str(sampleStDev))
		print('Population standard deviation = ' + str(populationStDev))
		print('n = ' + str(n))
		print('Min = ' + str(minVal))
		print('Q1 = ' + str(q1Val))
		print('Median = ' + str(medianVal))
		print('Q3 = ' + str(q3Val))
		print('Max = ' + str(maxVal))
		print('Sum of squared deviations = ' + str(ssx))
	except FileNotFoundError:
		print('File not found')
	except KeyError:
		print('Sheet does not exist')
	except TypeError:
		print('File contains letters or empty cells')

#NOTE: Should output mean, sum of x, sum of x-squared, stDev, Population stDev, n, five number summary for both variables, sum of standard deviations for both variables
def twoVarStats(src, setSheet, column1, column2):
