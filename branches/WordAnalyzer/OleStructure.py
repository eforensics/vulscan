#!/usr/bin/python
# -*- coding: utf-8 -*-

'''
author: Computists
website: http://computists.tistory.com
last modified: Feb 2013
'''
from struct import *
from collections import namedtuple

#== Defined ==#

SECID_FREE	 = -1
SECID_END	 = -2
SECID_SAT	 = -3
SECID_MSAT	 = -4
SIZE_DIRENTRY = 128
SIZE_HEADER = 512
ERROR	= -1

#======================================#

#== OLE Header Structure ==#

ruleChar_Header = '8s16s5h10s8l436s'
ruleName_Header = ('Sig UID Revision Version ByteOrder SSZ SSSZ Reserved SecCount_SAT SecID_DirStart signature MiniSectorCutoff SecID_SSATStart SecCount_SSAT SecID_MSATStart SecCount_MSAT MSAT')

#======================================#

#== OLE Directory Entry Structure ==#

ruleChar_DirEntry = '64sH2B3l16sl8s8s3l'
ruleName_DirEntry = ('DirName Size Type Color LeftChild RightChild RootChild UID Flags createTime LastModifyTime SecID_Start StreamSize Reserved')

#======================================#


#======================================#

#== Defined by header data ==#

def followSecs (SATTable,num,totalData, SECSIZE):
	Secs = []
	SecID = num
	SecCount = 0
	Sector_Size = SECSIZE
	offset = SIZE_HEADER
	ss =[]

	while(True):
		ss.append(SecID)
		SecCount += 1
		pSec = offset+SecID*Sector_Size
		Secs.append(totalData[pSec:pSec+Sector_Size])
		try: ### 0b3aac1df3ef1631f2539f51a1bdeb54.bin
			SecID = SATTable[SecID]
		except:
			print"\t[!] Failed to get SecID | Sector Numbers in SAT are out of range"
			print"\t\t[-] Table Length : %d"%len(SATTable)
			print"\t\t[-] Access point : %s"%str(SecID)
			return False, False
		
		if SecID == SECID_END :
			break
	retSecs = ''.join(Secs)
	return retSecs,SecCount

def followSSecs (SSATTable,num,totalData, SSECSIZE):
	if num == -2:
		return '', 0

	Secs = []
	SecID = num
	SecCount = 0
	Sector_Size = SSECSIZE
	offset = 0
	ss = []	

	while(True):
		ss.append(SecID)
		SecCount += 1
		pSec = offset+SecID*Sector_Size
		Secs.append(totalData[pSec:pSec+Sector_Size])
		try: ### 3ff16fa9ba2968f86d87bc5c092b2dc1.bin
			SecID = SSATTable[SecID]
		except:
			print"\t[!] Failed to get SecID | Sector Numbers in SSAT are out of range"
			print"\t\t[-] Table Length : %d"%len(SSATTable)
			print"\t\t[-] Access point : %s"%str(SecID)
			return False, False
		if SecID == SECID_END :
			break

	retSecs = ''.join(Secs)
	return retSecs,SecCount

def getSAT (totalData,SecCount_SAT, MSATTable, SECSIZE): ### SecCount_SAT == sHeader.SecCount_SAT
	loopcount = SecCount_SAT
	SAT = ""
	pSATs = []
	unpackMSAT = MSATTable

	for i in range(loopcount):
		pSAT = SIZE_HEADER + unpackMSAT[i]*SECSIZE
		SAT += totalData[pSAT:pSAT+SECSIZE]
		pSATs.append(pSAT)

	RuleSAT = '%dl'%((SecCount_SAT) * SECSIZE/4)

	try: ### 1cd6346e1709cad01e3c0ea190eecb27.bin
		SATTable = unpack(RuleSAT,SAT)
	except:
		print"\t[!] Failed to get SAT | Sector Numbers in SAT are out of range"
		print"\t\t[-] File Length : %d"%len(totalData)
		print"\t\t[-] Access point : %s"%str(pSATs)
		return False

	return SATTable

def getSSAT (totalData,SecID_SSAT, SecCount_SSAT, SATTable, SECSIZE): ### SecID_SSAT == sHeader.SecID_SSATStart | SecCount_SSAT == sHeader.SecCount_SSAT
	SSAT, countSSAT = followSecs (SATTable,SecID_SSAT,totalData, SECSIZE)
	if SSAT == False and countSSAT == False:
		return False
	
	RuleSSAT = '%dl'%((SecCount_SSAT) * SECSIZE/4)
	SSATTable = unpack(RuleSSAT,SSAT)

	return SSATTable

#======================================#

#== normal Func ==#

def multiTochr(data) :
	ret_data = ""
	for i in range(0,len(data),2) :
		ret_data += chr(ord(data[i:i+1]))
	return ret_data


def saveData (filename, data):
	try:
		f = open(filename,'wb')
		f.write(data)
		f.close()
	except:
		print "file write error"
	

#======================================#


def getMSAT (totalData, MSAT, SecCount_MSAT, SecID_MSATStart):
	MSATTable = unpack('109l',MSAT) ### default MSAT Table(109 tables)

	if SecCount_MSAT == 0: ### No additional MSAT tables
		return MSATTable
	
	else: ### exist MSAT sector tables 
		pNextMSATTable = 0
		pStartMSATTable = SecID_MSATStart
		pEnd = 0
		pStart = 0
		
		for i in range(SecCount_MSAT) :	
			
			pStart = (pStartMSATTable + 1) * 512
			pEnd = pStart + 512
			MSATData = totalData[pStart:pEnd]
			totalMSATTable = unpack('<128l',MSATData)
			MSATTable += totalMSATTable[:126]

			pStartMSATTable = totalMSATTable[127] ### Last SecID points to the next MSAT table

			if pStartMSATTable == -2 :
				break
	
	return MSATTable
