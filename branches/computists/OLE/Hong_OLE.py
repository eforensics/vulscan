# -*- coding: utf-8 -*-

import os, sys
from struct import *
from collections import namedtuple

#== Defined ==#

SECID_FREE	 = -1
SECID_END	 = -2
SECID_SAT	 = -3
SECID_MSAT	 = -4
SIZE_DIRENTRY = 128
SIZE_HEADER = 512

#======================================#

#== OLE Header Structure ==#

ruleChar_Header = '8s16s5h10s8l436s'
ruleName_Header = ('Sig UID Revision Version ByteOrder SSZ SSSZ Reserved SecCount_SAT SecID_DirStart signature MiniSectorCutoff SecID_SSATStart SecCount_SSAT SecID_MATStart SecCount_MAT MAT')

#======================================#

#== OLE Directory Entry Structure ==#

ruleChar_DirEntry = '64sH2B3l16sl8s8s3l'
ruleName_DirEntry = ('DirName Size Type Color LeftChild RightChild RootChild UID Flags createTime LastModifyTime SecID_Start StreamSize Reserved')

#======================================#

#== get a ole header information ==#

writeFileName = "StreamData.txt"
hwFile = open(writeFileName,'wb')

#filename = 'hwp_test.hwp'
filename = sys.argv[1]
hFile = open(filename,'rb')
totalData = hFile.read()
hFile.close()
data = totalData[0:SIZE_HEADER]

ntHeader = namedtuple('Header',ruleName_Header)
sHeader = ntHeader._make(unpack(ruleChar_Header,data))

print(sHeader)###
hwFile.write("Header\n"+str(sHeader)+"\n\n")

#======================================#

#== Defined by header data ==#

SECSIZE = 2 ** sHeader.SSZ
SSECSIZE = 2 ** sHeader.SSSZ
MINSECSIZE = sHeader.MiniSectorCutoff

def followSecs (SATable,num,totalData):
	Secs = ""
	SecID = num
	SecCount = 0
	Sector_Size = SECSIZE
	offset = SIZE_HEADER
	ss =[]

	while(True):
		ss.append(SecID)
		SecCount += 1
		pSec = offset+SecID*Sector_Size
		Secs += totalData[pSec:pSec+Sector_Size]
		SecID = SATable[SecID]
		if SecID == SECID_END :
			break
	print(str(ss))
	return Secs,SecCount

def followSSecs (SSATable,num,totalData):
	Secs = ""
	SecID = num
	SecCount = 0
	Sector_Size = SSECSIZE
	offset = 0
	ss = []	

	while(True):
		ss.append(SecID)
		SecCount += 1
		pSec = offset+SecID*Sector_Size
		Secs += totalData[pSec:pSec+Sector_Size]
		SecID = SSATable[SecID]
		if SecID == SECID_END :
			break
	print(str(ss))
	return Secs,SecCount

def getSAT (totalData,sHeader):
	if sHeader.SecCount_MAT == 0 :
		loopcount = sHeader.SecCount_SAT
		SAT = ""
		unpackMAT = unpack('109l',sHeader.MAT)

		for i in range(loopcount):
			pSAT = SIZE_HEADER + unpackMAT[i]*SECSIZE
			SAT += totalData[pSAT:pSAT+SECSIZE]
	else:
		None
	
	RuleSAT = '%dl'%((sHeader.SecCount_SAT) * SECSIZE/4)
	SATable = unpack(RuleSAT,SAT)

	return SATable

def getSSAT (totalData,sHeader):
	SecID_SSAT = sHeader.SecID_SSATStart
	SecCount_SSAT = sHeader.SecCount_SSAT

	SSAT, countSSAT = followSecs (SATable,SecID_SSAT,totalData)
	
	RuleSSAT = '%dl'%((sHeader.SecCount_SSAT) * SECSIZE/4)
	print(calcsize(RuleSSAT))
	SSATable = unpack(RuleSSAT,SSAT)

	return SSATable

#======================================#

#== get a Sector Allocation Table ==#

SATable = getSAT(totalData,sHeader)
print(SATable)###

#======================================#

#== get a Short-Sector Allocation Table ==#

SSATable = getSSAT(totalData,sHeader)
#print(SSATable)###

hwFile.write("SATable\n")
for i in range (len(SATable)) :
	hwFile.write("%d>%d\t"%(i,SATable[i]))
	if i%10 == 0 :
		hwFile.write("\n")		

hwFile.write("\n\n")

hwFile.write("SSATable\n")
for i in range (len(SSATable)) :
	hwFile.write("%d>%d\t"%(i,SSATable[i]))
	if i%10 == 0 :
		hwFile.write("\n")		
hwFile.write("\n\n")

#======================================#

#== get a Directory Entry information ==#

totalDir, DirCount = followSecs(SATable,sHeader.SecID_DirStart,totalData)
ntDirEntry = namedtuple('DirEntry',ruleName_DirEntry)
aDirEntry = []
her = 0
for i in range((DirCount*4)) :
	aDirEntry.append(ntDirEntry._make(unpack(ruleChar_DirEntry,totalDir[i*SIZE_DIRENTRY:SIZE_DIRENTRY+i*SIZE_DIRENTRY])))
	if aDirEntry[i-her].Size == 0 :
		her += 1
		aDirEntry.pop()

#print(aDirEntry)###
hwFile.write("Directory Entry\n"+str(aDirEntry)+"\n\n")

#======================================#

#== get Streams from each Directory ==#

for i in range(len(aDirEntry)) :
	if aDirEntry[i].Type == 5:
		offsetSSATdata = aDirEntry[i].SecID_Start
		break

SSecData, SSecCount = followSecs(SATable,offsetSSATdata,totalData)

for i in range(len(aDirEntry)) :
	hwFile.write("DirName : "+aDirEntry[i].DirName+"\n")
	if aDirEntry[i].StreamSize >= MINSECSIZE or aDirEntry[i].Type == 5:
		dirData, dirSecCount = followSecs(SATable,aDirEntry[i].SecID_Start,totalData)
	else:
		dirData, dirSecCount = followSSecs(SSATable,aDirEntry[i].SecID_Start,SSecData)
	hwFile.write(dirData+"\n\n")
	if i == 6 :
		getFIB(dirData)
	

#== get FIB Structure ==#

def getFIB (Data):
ruleChar_FIB = '8s16s5h10s8l436s'
ruleName_FIB = ('Sig UID Revision Version ByteOrder SSZ SSSZ Reserved SecCount_SAT SecID_DirStart signature MiniSectorCutoff SecID_SSATStart SecCount_SSAT SecID_MATStart SecCount_MAT MAT')
ntFIB = namedtuple('Header',ruleName_FIB)
sFIB = ntHeader._make(unpack(ruleChar_FIB,Data[:))

#======================================#

#== get FIB Structure ==#

def getFibBase (Data):
ruleChar_FibBase = '5hhhlbbhhll"
ruleName_FibBase = ('wIdent nFib unused lid pnNext opt1 nFibBack lkey envr opt2 reserved3 reserved4 reserved5 reserved6')
ntFIB = namedtuple('Header',ruleName_FibBase)
sFIB = ntHeader._make(unpack(ruleChar_FibBase,Data[:))

#======================================#

#== get FIB Structure ==#

def getFIB (dirData):
ruleChar_FIB = '8s16s5h10s8l436s'
ruleName_FIB = ('Sig UID Revision Version ByteOrder SSZ SSSZ Reserved SecCount_SAT SecID_DirStart signature MiniSectorCutoff SecID_SSATStart SecCount_SSAT SecID_MATStart SecCount_MAT MAT')
ntFIB = namedtuple('Header',ruleName_FIB)
sFIB = ntHeader._make(unpack(ruleChar_FIB,dirData[:))

#======================================#



hwFile.close()	

#======================================#

#== Output on screen ==#

print "%s\t\t\t%s\t%s\t%s\t%s\t%s\t%s"%("Directory name", "Type", "¢Ð.C", "¢Ñ.C", "R.C", "SecID", "Size")
for i in range(len(aDirEntry)) :
	print "%d %s\t%d\t%d\t%d\t%d\t%d\t%d"%(i, aDirEntry[i].DirName[:28], aDirEntry[i].Type ,aDirEntry[i].LeftChild, aDirEntry[i].RightChild, aDirEntry[i].RootChild, aDirEntry[i].SecID_Start,aDirEntry[i].StreamSize)

#======================================#
