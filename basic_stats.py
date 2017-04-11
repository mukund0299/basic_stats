#TODO: Test all functions

#IDEA: Create a GUI for the user rather than CLI

import openpyxl
import mpmath
import statistics
import matplotlib.pyplot as plt
import pylab

from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile


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
			result = 0.5*mpmath.betainc((df/2), 0.5, 0, x2, regularized=True)
		else:
			x2 = (i**2)/((i**2)+df)
			result = 0.5*(mpmath.betainc(0.5, (df/2), 0, x2, regularized=True)+1)
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


class OneSampleZInterval(object):
	def __init__(self, master):
		#Setting up the main window and central frame
		top = self.top = Toplevel(master)
		self.frame = Frame(top)
		self.frame.grid(column=0, row=0, sticky=(N, W, E, S))
		self.frame.columnconfigure(0, weight=1)
		self.frame.rowconfigure(0, weight=1)
		#Creating Labels for the inputs
		self.successesLabel = Label(self.frame, text='Enter number of successes')
		self.nLabel = Label(self.frame, text='Enter size of sample')
		self.confLabel = Label(self.frame, text='Enter confidence level (0-100)')
		self.successesLabel.grid(row=0, sticky=E)
		self.nLabel.grid(row=1, sticky=E)
		self.confLabel.grid(row=2, sticky=E)
		#Creating input dialogs and gridding them
		self.successesInput = Entry(self.frame)
		self.nInput = Entry(self.frame)
		self.confInput = Entry(self.frame)
		self.successesInput.grid(row=0, column=1)
		self.nInput.grid(row=1, column=1)
		self.confInput.grid(row=2, column=1)
		#Creating Button to retrieve all values
		self.complete = Button(self.frame, text='Submit', command=self.cleanup)
		self.complete.grid(columnspan=2)

	def cleanup(self):
		self.successesVal = int(self.successesInput.get())
		self.nVal = int(self.nInput.get())
		self.confVal = int(self.confInput.get())
		self.oneSampleZInterval(self.successesVal, self.nVal, self.confVal)
		self.top.destroy()

	def oneSampleZInterval(self, successes, n, confidenceLevel):
		#Calculating p from the sample
		self.pHat = successes/n
		#Getting Critical value
		self.zCritical = -1*invNorm((1-(confidenceLevel/100))/2)
		#Getting Standard error
		self.stDev = float(mpmath.sqrt(((self.pHat)*(1-self.pHat))/(n)))
		self.marError = self.stDev*self.zCritical
		self.CLower = self.pHat - self.marError
		self.CUpper = self.pHat + self.marError
		self.output = ('p-Hat = ' + str(self.pHat) + '\n' + 'ME = ' + str(self.marError) + '\n' + 'CLower = ' + str(self.CLower) + '\n' + 'CUpper = ' + str(self.CUpper))


class TwoSampleZInterval(object):
	def __init__(self, master):
		#Setting up the main window and central frame
		top = self.top = Toplevel(master)
		self.frame = Frame(top)
		self.frame.grid(column=0, row=0, sticky=(N, W, E, S))
		self.frame.columnconfigure(0, weight=1)
		self.frame.rowconfigure(0, weight=1)
		#Creating Labels for the inputs
		self.successesLabel1 = Label(self.frame, text='Enter number of successes for sample 1')
		self.nLabel1 = Label(self.frame, text='Enter size of sample 1')
		self.successesLabel2 = Label(self.frame, text='Enter number of successes for sample 2')
		self.nLabel2 = Label(self.frame, text='Enter size of sample 2')
		self.confLabel = Label(self.frame, text='Enter confidence level (0-100)')
		self.successesLabel1.grid(row=0, sticky=E)
		self.nLabel1.grid(row=1, sticky=E)
		self.successesLabel2.grid(row=2, sticky=E)
		self.nLabel2.grid(row=3, sticky=E)
		self.confLabel.grid(row=4, sticky=E)
		#Creating input dialogs and gridding them
		self.successesInput1 = Entry(self.frame)
		self.nInput1 = Entry(self.frame)
		self.successesInput2 = Entry(self.frame)
		self.nInput2 = Entry(self.frame)
		self.confInput = Entry(self.frame)
		self.successesInput1.grid(row=0, column=1)
		self.nInput1.grid(row=1, column=1)
		self.successesInput2.grid(row=2, column=1)
		self.nInput2.grid(row=3, column=1)
		self.confInput.grid(row=4, column=1)
		#Creating Button to retrieve all values
		self.complete = Button(self.frame, text='Submit', command=self.cleanup)
		self.complete.grid(row=5, column=1)

	def cleanup(self):
		self.successesVal1 = int(self.successesInput1.get())
		self.nVal1 = int(self.nInput1.get())
		self.successesVal2 = int(self.successesInput2.get())
		self.nVal2 = int(self.nInput2.get())
		self.confVal = int(self.confInput.get())
		self.twoSampleZInterval(self.successesVal1, self.nVal1, self.successesVal2, self.nVal2, self.confVal)
		self.top.destroy()

	def twoSampleZInterval(self, successes1, n1, successes2, n2, confidenceLevel):
		#Calculating p from the samples
		self.pHat1 = successes1/n1
		self.pHat2 = successes2/n2
		#Getting the difference in p
		self.pDiff = self.pHat1-self.pHat2
		#Getting critical value
		self.zCritical = -1*invNorm((1-(confidenceLevel/100))/2)
		#Getting Standard error
		self.stDev = mpmath.sqrt((((self.pHat1)*(1-self.pHat1))/n1) + (((self.pHat2)*(1-self.pHat2))/n2))
		self.marError = self.zCritical*self.stDev
		self.CLower = self.pDiff - self.marError
		self.CUpper = self.pDiff + self.marError
		self.output = ('p-Hat1 = ' + str(self.pHat1) + '\n' + 'pHat2 = ' + str(self.pHat2) + '\n' + 'pDiff = ' + str(self.pDiff) + '\n' + 'ME  = ' + str(self.pDiff) + '\n' + 'Clower = ' + str(self.CLower) + '\n' + 'CUpper = ' + str(self.CUpper))


#Confidence Intervals for means; reading from an excel file
def oneSampleTInterval(src, setSheet, confidenceLevel, column):
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
		tCritical = -1*invt(((1-(confidenceLevel/100))/2), df)
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
		#Getting mean, Standard Deviation, n and df for both samples
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
		df = (n1+n2)-2
		stDev = float(mpmath.sqrt(((sampleStDev1**2)/n1)+((sampleStDev2**2)/n2)))
		#Getting critical value
		tCritical = -1*invt(((1-(confidenceLevel/100))/2), df)
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


class OneVarStats(object):
	def __init__(self, master):
		#Setting up the main window and central frame
		top = self.top = Toplevel(master)
		self.frame = Frame(top)
		self.frame.grid(column=0, row=0, sticky=(N, W, E, S))
		self.frame.columnconfigure(0, weight=1)
		self.frame.rowconfigure(0, weight=1)
		#Creating Labels for the inputs
		self.sheetLabel = Label(self.frame, text='Enter the name of the sheet')
		self.columnLabel = Label(self.frame, text='Enter column number')
		self.sheetLabel.grid(row=0, sticky=E)
		self.columnLabel.grid(row=1, sticky=E)
		#Creating inputs
		self.srcVAL = filedialog.askopenfilename(title="Locate Excel file", filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
		self.sheetVal = Entry(self.frame)
		self.sheetVal.grid(row=0, column=1)
		self.columnVal = Entry(self.frame)
		self.columnVal.grid(row=1, column=1)
		#Creating Button to retrieve all values
		self.complete = Button(self.frame, text='Submit', command=self.cleanup)
		self.complete.grid(row=2, column=1)

	def cleanup(self):
		self.source = self.srcVAL
		self.sheetName = self.sheetVal
		self.column = int(self.columnVal.get())
		self.oneVarStats(self.source, self.sheetName, self.column)
		self.top.destroy()

	def oneVarStats(src, setSheet, column):
		try:
			#Opening the workbook and setting the sheet
			self.wb = openpyxl.load_workbook(src)
			self.sheet = wb.get_sheet_by_name(setSheet)
			#Getting a list of the cells in order to calculate length
			self.dataList = list(self.sheet.columns)[column]
			#Populating a list with data from the cells
			self.data = []
			for cellObj in self.dataList:
				self.data.append(float(cellObj.value))
			self.mean = statistics.mean(self.data)
			self.sampleStDev = statistics.stdev(self.data)
			self.populationStDev = statistics.pstdev(self.data)
			self.n = len(self.data)
			self.sumOfValues = sumData(self.data, 1)
			self.sumOfValuesSquared = sumData(selfdata, 2)
			#Five Number summary
			self.data.sort()
			self.minVal = self.data[0]
			self.q1Val, self.medianVal, self.q3Val = 0, 0, 0
			self.maxVal = self.data[len(self.data)-1]
			#If the data has odd number of values
			if (len(self.data) % 2 != 0):
				self.middleIndex = (len(self.data)-1)//2
				self.medianVal = self.data[self.middleIndex]
				self.lowerData = self.data[0:self.middleIndex]
				self.q1Val = statistics.median(self.lowerData)
				self.upperData = data[self.middleIndex+1:len(self.data)]
				self.q3Val = statistics.median(self.upperData)
			else:
				self.medianVal = statistics.median(self.data)
				self.lowerData = self.data[0:len(self.data)//2]
				self.q1Val = statistics.median(self.lowerData)
				self.upperData = self.data[len(self.data)//2:len(self.data)]
				self.q3Val = statistics.median(self.upperData)
			#Sum of squared deviations
			self.ssx = 0
			for value in self.data:
				self.ssx += (value-self.mean)**2
			#Histogram and Boxplot
			self.fig = plt.figure()
			self.ax = self.fig.add_subplot(111)
			self.ax.hist(self.data)
			self.fig2 = plt.figure()
			self.ax2 = self.fig2.add_subplot(111)
			self.ax2.boxplot(self.data)
			self.output = ('Mean = ' + str(mean))
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
			plt.show()
		except FileNotFoundError:
			print('File not found')
		except KeyError:
			print('Sheet does not exist')
		except TypeError:
			print('File contains letters or empty cells')


def linReg(src, setSheet, column1, column2):
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
		if (len(data1) != len(data2)):
			raise Exception('Data is not of equal n')
		(m, b) = pylab.polyfit(data1, data2, 1)
		yp = pylab.polyval([m, b], data1)
		mean1 = statistics.mean(data1)
		mean2 = statistics.mean(data2)
		stDev1 = statistics.stdev(data1)
		stDev2 = statistics.stdev(data2)
		r = (m*stDev1)/stDev2
		rSquared = r**2
		print('a = ' + str(b))
		print('b = ' + str(m))
		print('r = ' + str(r))
		print('r-squared = ' + str(rSquared))
		fig = plt.figure()
		ax = fig.add_subplot(111)
		ax.plot(data1, data2, 'ro')
		ax.plot(data1, yp)
		plt.show()
	except FileNotFoundError:
		print('File not found')
	except KeyError:
		print('Sheet does not exist')
	except TypeError:
		print('File contains letters or empty cells')
	except Exception as inst:
		print(inst)


class main(object):
	def __init__(self, master):
		self.master = master
		# *******Menu Bar********
		self.menubar = Menu(master)
		self.master.config(menu=self.menubar)
		#Adding stats calculations sub-menu
		self.statCalc = Menu(self.menubar)
		self.menubar.add_cascade(label='Stat Calculations', menu=self.statCalc)
		self.statCalc.add_command(label='One Variable Statistics', command=self.initOneVarStats)
		# self.statCalc.add_command(label='Linear Regression', command=initLinReg)
		#Adding confidence intervals sub-menu
		self.confInt = Menu(self.menubar)
		self.menubar.add_cascade(label='Confidence Intervals', menu=self.confInt)
		self.confInt.add_command(label='One Sample Z Interval', command=self.initOneSampleZInterval)
		self.confInt.add_command(label='Two Sample Z Interval', command=self.initTwoSampleZInterval)
		self.confInt.add_separator()
		# self.confInt.add_command(label='One Sample T Interval', command=initOneSampleTTest)
		# self.confInt.add_command(label='Two Sample T Interval', command=initTwoSampleTTest)
		# *******output area********
		self.resultLabel = Label(master, text='Output')
		self.result = Text(master, state='disabled')
		self.result.insert('1.0', 'Result will be shown here')
		self.resultLabel.pack()
		self.result.pack()

	def initOneSampleZInterval(self):
		self.w = OneSampleZInterval(self.master)
		self.master.wait_window(self.w.top)
		self.result['state'] = 'normal'
		self.result.delete('1.0', END)
		self.result.insert('1.0', self.w.output)
		self.result['state'] = 'disabled'

	def initTwoSampleZInterval(self):
		self.w = TwoSampleZInterval(self.master)
		self.master.wait_window(self.w.top)
		self.result['state'] = 'normal'
		self.result.delete('1.0', END)
		self.result.insert('1.0', self.w.output)
		self.result['state'] = 'disabled'

	def initOneVarStats(self):
		self.w = OneVarStats(self.master)
		self.master.wait_window(self.w.top)
		self.result['state'] = 'normal'
		self.result.delete('1.0', END)
		self.result.insert('1.0', self.w.output)
		self.result['state'] = 'disabled'


# *******Main Window********
root = Tk()
root.title('Basic Statistics')
root.geometry('300x500')
m = main(root)
root.mainloop()
