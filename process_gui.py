#!/usr/bin/env python
import warnings
warnings.simplefilter('ignore')

import os, re, string
import xlrd
from scipy import stats
import numpy
import wx
from Measurement import Measurement

# ============================================================================
# parameters
# ============================================================================

class Row:
	def __init__(self, name, values):
		self.name = name
		self.values = values

	def __str__(self):
		s = self.name + ': '
		for value in self.values:
			s += str(value) + ', '
		return s

	def getMeasurement(self, window = [1, 4], median = False):
		(start, end) = window
		values = getFloats(self.values[start:end+1])
		if median == True:
			return Measurement(numpy.median(values), numpy.std(values, ddof=1))
		else:
			return Measurement(numpy.mean(values), numpy.std(values, ddof=1))
			
	def getMaxWindow(self, windowSpan = 4, median = False):
		maxValue = 0
		window = [0, windowSpan - 1]
		for start in range(0, len(self.values) - windowSpan):
			end = start + windowSpan - 1
			values = getFloats(self.values[start:end+1])
			if (median == False):
				value = numpy.mean(values)
			else:
				value = numpy.median(values)
			if (value > maxValue):
				maxValue = value
				window = [start, end]
		return window

class RowGroup:
	def __init__(self, name):
		self.name = name
		self.rows = []

	def __str__(self):
		nRows = len(self.rows)
		s = 'group name = ' + self.name + ' (' + str(nRows) + ')'
		for row in self.rows:
			s += '\n' + str(row)
		return s + '\n'
	
	def getMeasurements(self, window = [1, 4], median = False):
		measurements = []
		for row in self.rows:
			measurement = row.getMeasurement(window, median)
			measurements.append(measurement)
		return measurements
	
	def getMaxMeasurements(self, windowSpan = 4, median = False):
		measurements = []
		for row in self.rows:
			window = row.getMaxWindow(windowSpan, median)
			measurement = row.getMeasurement(window, median)
			measurements.append(measurement)
		return measurements

def getMean(measurements):
	(valueSum, errorSum, n) = (0, 0, 0)
	for measurement in measurements:
		valueSum += measurement.value
		errorSum += measurement.error
		n = n + 1
	return Measurement(valueSum / n, errorSum / n)
	
def getFloats(inputValues):
	values = []
	for value in inputValues:
		if isinstance(value, float):
			values.append(value)
	return values

def getMedian(measurements):
	values = getValues(measurements)
	return numpy.median(values)
	
def getValues(measurements):
	values = []
	for measurement in measurements:
		values.append(measurement.value)
	return values

def readFile(filename, useSpecificGroupName = False):
	try:
		wb = xlrd.open_workbook(filename)
	except Exception, e:
		print 'skipped: file could not be opened by xlrd library'
		return

	try:
		dataSheet = wb.sheet_by_name('Data')
	except Exception, e:
		print 'skipped: file does not contain "Data" worksheet' 
		return

	# create RowGroups from frows in excel sheet
	rowGroups = {}
	for iRow in range(dataSheet.nrows):
		data = []
		for iCol in range(dataSheet.ncols):
			cell = dataSheet.cell_value(iRow, iCol)
			data.append(cell)	
		rowName = data[0];
		rowValues = data[1:]
		row = Row(rowName, rowValues)

		# get group name from string 'name-id-number'
		if (useSpecificGroupName):
			pattern = r'^(\w+(-| )?\d+)-\d+$' # name-id
		else:
			pattern = r'^(\w+)(-| )?\d+-\d+$' # name
		match = re.match(pattern, rowName)
		if match == None:
			print 'could not get group name from row: ', rowName
		else:			
			matches = match.groups()
			groupName = matches[0]
			if (rowGroups.has_key(groupName) == False):
				rowGroups[groupName] = RowGroup(groupName)
			rowGroups[groupName].rows.append(row)

	return rowGroups

# ============================================================================
# main
# ============================================================================

class AppFrame(wx.Frame):
	def __init__(self, *args, **kwds):
		# begin wxGlade: AppFrame.__init__
		kwds["style"] = wx.DEFAULT_FRAME_STYLE
		wx.Frame.__init__(self, *args, **kwds)
		self.directoryLabel = wx.StaticText(self, -1, "source folder: ")
		self.directoryText = wx.TextCtrl(self, -1, "")
		self.directoryText.SetEditable(False)
		self.browseButton = wx.Button(self, 20, "browse...")
		self.rowLabel = wx.StaticText(self, -1, "row calculation: ")
		self.rowMeanRadio = wx.RadioButton(self, 30, "mean", style=wx.RB_GROUP )
		self.rowMedianRadio = wx.RadioButton(self, 40, "median")
		self.windowRadio = wx.RadioButton(self, 50, "window (min / max): ", style=wx.RB_GROUP)
		self.windowMinText = wx.TextCtrl(self, 60, "2")
		self.windowMaxText = wx.TextCtrl(self, 70, "5")
		self.maxWindowRadio = wx.RadioButton(self, -1, "max window (span):")
		self.windowSpanText = wx.TextCtrl(self, 80, "4")
		self.analyzeButton = wx.Button(self, 90, "analyze")
		self.resultLabel = wx.StaticText(self, -1, "")
		
		self.__set_properties()
		self.__do_layout()
		
		wx.EVT_BUTTON(self, 20, self.OnBrowseClick)
		wx.EVT_BUTTON(self, 90, self.OnAnalyzeClick)

		# end wxGlade

	def __set_properties(self):
		# begin wxGlade: AppFrame.__set_properties
		self.SetTitle("Joe's Program")
		self.directoryText.SetMinSize((300, 27))
		self.windowRadio.SetValue(1)
		# end wxGlade

	def __do_layout(self):
		# begin wxGlade: AppFrame.__do_layout
		sizer_1 = wx.BoxSizer(wx.VERTICAL)
		sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
		sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
		sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
		sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
		sizer_6 = wx.BoxSizer(wx.HORIZONTAL)
		sizer_2.Add(self.directoryLabel, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_2.Add(self.directoryText, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_2.Add(self.browseButton, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_1.Add(sizer_2, 1, wx.EXPAND, 0)
		sizer_3.Add(self.rowLabel, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_3.Add(self.rowMeanRadio, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_3.Add(self.rowMedianRadio, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_1.Add(sizer_3, 1, wx.EXPAND, 0)
		sizer_4.Add((50, 20), 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_4.Add(self.windowRadio, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_4.Add(self.windowMinText, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_4.Add(self.windowMaxText, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_1.Add(sizer_4, 1, wx.EXPAND, 0)
		sizer_5.Add((50, 20), 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_5.Add(self.maxWindowRadio, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_5.Add(self.windowSpanText, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_1.Add(sizer_5, 1, wx.EXPAND, 0)
		sizer_1.Add((50, 20), 0, wx.ALIGN_CENTER_VERTICAL, 0)
		
		sizer_6.Add((50, 20), 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_6.Add(self.analyzeButton, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_6.Add(self.resultLabel, 0, wx.ALIGN_CENTER_VERTICAL, 0)
		sizer_1.Add(sizer_6, 1, wx.EXPAND, 0)
		self.SetSizer(sizer_1)
		sizer_1.Fit(self)
		self.Layout()
		# end wxGlade

	def OnBrowseClick(self, event):
		dialog = wx.DirDialog (self, style = wx.DD_DIR_MUST_EXIST )
		if dialog.ShowModal() == wx.ID_OK:
		   self.directoryText.ChangeValue(dialog.GetPath())

	def OnAnalyzeClick(self, event): 
		# script perfoms analysis on each excel file in the given directory
		directory = self.directoryText.GetValue()
		
		# time points to use in a replicate's value. if Nothing, the script finds the
		# window giving the greatest value
		start = int(self.windowMinText.GetValue())
		end = int(self.windowMaxText.GetValue())
		window = [start, end]
		
		getMaxWindow = False
		if self.maxWindowRadio.GetValue() == True:
			getMaxWindow = True
		windowSpan = int(self.windowSpanText.GetValue())
		
		# decide whether to use mean or median in computations
		getReplicateMedian = False
		if self.rowMedianRadio.GetValue() == True:
			getReplicateMedian = True
		
		fout = open(directory + '/output.txt', 'w')
		for filename in os.listdir(directory):
			rowGroups = None
			if re.match(r'.+\.xls', filename):
				fout.write('processing: ' + filename + "\n")
				rowGroups = readFile(directory + '/' + filename)
				
			if rowGroups is not None:	
				# get group that matches 'WT' for reference
				refGroup = None
				for groupName, rowGroup in rowGroups.iteritems():
					search = groupName.lower();
					pattern = r'^wt' # name-id
					if (re.match(pattern, search)):
						refGroup = rowGroup
						
				if refGroup is None:
					print 'did not find reference group named "WT", t-tests not performed'
				else:
					refMeasurements = []
					if getMaxWindow:
						refMeasurements = refGroup.getMaxMeasurements(windowSpan, getReplicateMedian)
					else:
						refMeasurements = refGroup.getMeasurements(window, getReplicateMedian)
					refValues = getValues(refMeasurements)
					refMean = getMean(refMeasurements)
					refMedian = getMedian(refMeasurements)
				
				# for each group, do a t-test vs reference values
				for groupName, rowGroup in rowGroups.iteritems():
					testMeasurements = []
					if getMaxWindow:
						testMeasurements = rowGroup.getMaxMeasurements(windowSpan, getReplicateMedian)
					else:
						testMeasurements = rowGroup.getMeasurements(window, getReplicateMedian)
					testValues = getValues(testMeasurements)
					mean = getMean(testMeasurements)
					median = getMedian(testMeasurements)
					fout.write('name: ' + rowGroup.name + "\n")
					fout.write(str(testValues) + "\n")
					fout.write('mean = ' + str(mean) + ', median = ' + str(median) + "\n")
					
					if refGroup is not None:
						meanChange = (mean.value - refMean.value) / refMean.value * 100
						medianChange = (median - refMedian) / refMedian * 100
						fout.write('percent change: mean = ' + str(meanChange) + '%, median = ' + str(medianChange) + '%' + "\n")
		
						(t, p) = stats.ttest_ind(testValues, refValues)
						fout.write('2-tailed independent t-test vs. ' + refGroup.name + ':' + "\n")
						fout.write('t = ' + str(t) + ', p = ' + str(p) + "\n")
					fout.write("\n")
			fout.write("\n")
		self.resultLabel.SetLabel("wrote output.txt")
		

# end of class AppFrame

if __name__ == "__main__":
	app = wx.PySimpleApp(0)
	wx.InitAllImageHandlers()
	MyAppFrame = AppFrame(None, -1, "")
	app.SetTopWindow(MyAppFrame)
	MyAppFrame.Show()
	app.MainLoop()






