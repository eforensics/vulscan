# -*- coding: utf-8 -*-
from struct import *

#== bitParse func ==#
""" Usage :		bitParse(seperated bit count,Number)

	ex.1)		bitParse('1,1,1,1',0xc)
	result)		[0, 0, 1, 1]

	ex.2)		bitParse('1,1,?',0xc)
	result)		[0, 0, 3]
"""

def bitParse(ruleString, data) :
	retData = []
	bitCount = 0
	ruleString = ruleString.split(',')
	for i in ruleString :
		if i == "?": 
			i = findBitCount(data)
		else:
			i = int(i)
		retData.append((data / 2**bitCount) & (2**i-1))
		bitCount += int(i)
	return retData

def findBitCount (data):
	dCount = 0
	while(True) :
		data = data/2
		dCount += 1
		if data == 0:
			break
	return dCount

#== RecordHeader Structure ==#
# structure size : 22

ruleChar_RecordHeader = '10,10,12'
ruleName_RecordHeader = ('TagID Level Size')

#======================================#

#== RecordHeader Structure ==#
# structure size : 22

ruleChar_Record = '10,10,12'
ruleName_Record = ('Header Data')

#======================================#
