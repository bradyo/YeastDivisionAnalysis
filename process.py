#!/usr/bin/env python
import os, re, string
import xlrd, xlwt
from pprint import pprint
from scipy import stats
import numpy
import warnings
from Measurement import *
warnings.simplefilter('ignore')

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
	print 'processing: ', filename	
	try:
		wb = xlrd.open_workbook(filename)
	except Exception as e:
		print 'skipped: file could not be opened by xlrd library'
		return

	try:
		dataSheet = wb.sheet_by_name('Data')
	except Exception as e:
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

# script perfoms analysis on each excel file in the given directory
directory = os.getcwd() + '/input'

# time points to use in a replicate's value. if Nothing, the script finds the
# window giving the greatest value
window = [1, 4]
getMaxWindow = True
windowSpan = 4

# decide whether to use mean or median in computations
getReplicateMedian = True

for filename in os.listdir(directory):
	rowGroups = None
	if re.match(r'.+\.xls', filename):
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
			refMean = getMean(refMeasurements).value
			refMedian = getMedian(refMeasurements)
		
		# for each group, do a t-test vs reference values
		for groupName, rowGroup in rowGroups.iteritems():
			testMeasurements = []
			if getMaxWindow:
				testMeasurements = rowGroup.getMaxMeasurements(windowSpan, getReplicateMedian)
			else:
				testMeasurements = rowGroup.getMeasurements(window, getReplicateMedian)
			testValues = getValues(testMeasurements)
			mean = getMean(testMeasurements).value
			median = getMedian(testMeasurements)
			print 'name:', rowGroup.name
			print testValues
			print 'mean = ' + str(mean) + ', median = ' + str(median)
			
			if refGroup is not None:
				meanChange = (mean - refMean) / refMean * 100
				medianChange = (median - refMedian) / refMedian * 100
				print 'percent change: mean = ' + str(meanChange) + '%, median = ' + str(medianChange) + '%'

				(t, p) = stats.ttest_ind(testValues, refValues)
				print '2-tailed independent t-test vs. ' + refGroup.name + ': t = ' + str(t) + ', p = ' + str(p)
			print
	print
	print
