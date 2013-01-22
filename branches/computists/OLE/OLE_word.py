# -*- coding: utf-8 -*-

import sys, string
from WordStructure import *
from struct import *
from collections import namedtuple
from pprint import pprint


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

#print(sHeader)###
hwFile.write("Header\n"+str(sHeader)+"\n\n")

#======================================#

#== Defined by header data ==#

SECSIZE = 2 ** sHeader.SSZ
SSECSIZE = 2 ** sHeader.SSSZ
MINSECSIZE = sHeader.MiniSectorCutoff

def followSecs (SATTable,num,totalData):
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
		SecID = SATTable[SecID]
		if SecID == SECID_END :
			break
#	print(str(ss))
	retSecs = ''.join(Secs)
	return retSecs,SecCount

def followSSecs (SSATTable,num,totalData):
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
		SecID = SSATTable[SecID]
		if SecID == SECID_END :
			break
#	print(str(ss))
	retSecs = ''.join(Secs)
	return retSecs,SecCount

def getSAT (totalData,SecCount_SAT, MSATTable): ### SecCount_SAT == sHeader.SecCount_SAT
	loopcount = SecCount_SAT
	SAT = ""
	unpackMSAT = MSATTable

	for i in range(loopcount):
		pSAT = SIZE_HEADER + unpackMSAT[i]*SECSIZE
		SAT += totalData[pSAT:pSAT+SECSIZE]
	
	print "Count_SAT : %d" %sHeader.SecCount_SAT

	RuleSAT = '%dl'%((sHeader.SecCount_SAT) * SECSIZE/4)
	SATTable = unpack(RuleSAT,SAT)

	return SATTable

def getSSAT (totalData,SecID_SSAT, SecCount_SSAT): ### SecID_SSAT == sHeader.SecID_SSATStart | SecCount_SSAT == sHeader.SecCount_SSAT

	SSAT, countSSAT = followSecs (SATTable,SecID_SSAT,totalData)
	
	RuleSSAT = '%dl'%((sHeader.SecCount_SSAT) * SECSIZE/4)
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

#== get a Master Sector Allocation Table ==#

MSATTable = getMSAT(totalData,sHeader.MSAT, sHeader.SecCount_MSAT, sHeader.SecID_MSATStart)
print "done MSAT table"

#======================================#

#== get a Sector Allocation Table ==#

SATTable = getSAT(totalData,sHeader.SecCount_SAT, MSATTable)
print "done SAT table"

#======================================#

#== get a Short-Sector Allocation Table ==#

SSATTable = getSSAT(totalData, sHeader.SecID_SSATStart, sHeader.SecCount_SSAT)

hwFile.write("SATTable\n")
for i in range (len(SATTable)) :
	hwFile.write("%d>%d\t"%(i,SATTable[i]))
	if i%10 == 0 :
		hwFile.write("\n")		

hwFile.write("\n\n")

hwFile.write("SSATTable\n")
for i in range (len(SSATTable)) :
	hwFile.write("%d>%d\t"%(i,SSATTable[i]))
	if i%10 == 0 :
		hwFile.write("\n")		
hwFile.write("\n\n")

print "done SSAT table"


#== get a Directory Entry information ==#

totalDir, DirCount = followSecs(SATTable,sHeader.SecID_DirStart,totalData)
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

print "done Directory Entry"

#======================================#

#== get Streams from each Directory ==#

for i in range(len(aDirEntry)) :
	if aDirEntry[i].Type == 5:
		offsetSSATdata = aDirEntry[i].SecID_Start
		break

SSecData, SSecCount = followSecs(SATTable,offsetSSATdata,totalData)

streamData = []

for i in range(len(aDirEntry)) :
	hwFile.write("DirName : "+aDirEntry[i].DirName+"\n")
	if aDirEntry[i].StreamSize >= MINSECSIZE or aDirEntry[i].Type == 5:
		dirData, dirSecCount = followSecs(SATTable,aDirEntry[i].SecID_Start,totalData)
	else:
		dirData, dirSecCount = followSSecs(SSATTable,aDirEntry[i].SecID_Start,SSecData)

	streamData.append(dirData)	
	hwFile.write(dirData+"\n\n")

hwFile.close()	

print "done follwing Sec"

#======================================#

#== Output on screen ==#

print "%s\t\t\t%s\t%s\t%s\t%s\t%s\t%s"%("Directory name", "Type", "☜.C", "☞.C", "R.C", "SecID", "Size")
for i in range(len(aDirEntry)) :
	print "%d %s\t%d\t%d\t%d\t%d\t%d\t%d"%(i, aDirEntry[i].DirName[:28], aDirEntry[i].Type ,aDirEntry[i].LeftChild, aDirEntry[i].RightChild, aDirEntry[i].RootChild, aDirEntry[i].SecID_Start,aDirEntry[i].StreamSize)

#======================================#

for i in range(len(aDirEntry)) : ## DirName trans
	DirName = multiTochr(aDirEntry[i].DirName)
	if DirName.find("WordDocument") != -1:
		wordStreamNum = i
	if DirName.find("1Table") != -1:
		Table1Num = i	
	if DirName.find("0Table") != -1:
		Table0Num = i

#########################

def saveData (fileName, data):
	f = open(fileName, 'wb');
	f.write(str(data))
	f.close()

#== Word file struct parsing ==#

sFIB = getFIB(streamData[wordStreamNum])
pprint(sFIB, open('FIB.txt','wb'))

saveData('WordStream.txt', streamData[wordStreamNum])

ntFibBase = namedtuple('FibBase',ruleName_FibBase)
sFibBase = ntFibBase._make(unpack(ruleChar_FibBase,sFIB.base))

ntFibBase_opt = namedtuple('FibBaseOpt',ruleName_FibBaseOpt)
sFibBase_opt = ntFibBase_opt._make(bitParse(ruleChar_FibBaseOpt,sFibBase.opt))

print "fWhichTblStm : %d" %sFibBase_opt.fWhichTblStm

sFibRgCswNew = getFibRgCswNew(sFIB.cswNew, sFIB.fibRgCswNew)

### sFibBase_opt.fWhichTblStm 값을 사용하여 Table 스트림의 스트림 넘버를 지정한다. 실질적으로 테이블을 설정하는 부분
if sFibBase_opt.fWhichTblStm == 1 :
	tableNum = Table1Num
else:
	tableNum = Table0Num
TableStream = streamData[tableNum]


sFibRgFcLcbBlob = getFibRgFcLcbBlob(sFIB.cbRgFcLcb ,sFIB.FibRgFcLcbBlob)

dataPlcBteChpx = TableStream[sFibRgFcLcbBlob.FibRgFcLcb97.fcPlcfBteChpx:sFibRgFcLcbBlob.FibRgFcLcb97.fcPlcfBteChpx+sFibRgFcLcbBlob.FibRgFcLcb97.lcbPlcfBteChpx]
sPlcBteChpx = getPlcBteChpx(dataPlcBteChpx)

### aFC 의 개수 만큼 sPnFkpChpx 가 존재하기 때문에 loop 문으로 모든 sPnFkpChpx 를 검색할 수 있도록 해야 한다.!!! ###

def aFC_Count(aFC):
	for i in aFC :
		yield i

aFC = aFC_Count(sPlcBteChpx.aFC)

for i in sPlcBteChpx.aPnBteChpx :
	pStart_ChpxFkp = i*512

	dataChpxFkp = streamData[wordStreamNum][pStart_ChpxFkp:pStart_ChpxFkp+512]
	ntChpxFkp = namedtuple('ChpxFkp',ruleName_ChpxFkp)
	sChpxFkp = ntChpxFkp._make(makeChpxFkp(dataChpxFkp))

	#======================================#

	#== get the information of Chpx structure ==#

	dataRgfc, dataRgb = getDataChpx(streamData[wordStreamNum],sChpxFkp,pStart_ChpxFkp)

	#======================================#

	"""
	#== #print dataRgb ==#

	for i in range(len(dataRgb)) :
		Pri,sprm = makePri(dataRgb[i][1])
		print("# [%d] #==========================================#"%i)
		print(Pri)	

	#======================================#

	count = 0
	for i in dataRgfc :
		print("[%d]"%count+" : "+i)
		count += 1
	"""

print "fEncrypted : %d"%sFibBase_opt.fEncrypted
print "fObfuscated : %d"%sFibBase_opt.fObfuscated

### check nFib ###
if sFIB.cswNew == 0 : 
	nFib = sFibBase.nFib
else:
	nFib = sFibRgCswNew.nFibNew

print "nFib value : %04X" %nFib
