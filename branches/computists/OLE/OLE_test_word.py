# -*- coding: utf-8 -*-

import sys
from struct import *
from collections import namedtuple
from WordStructure import *

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

#print(sHeader)###
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

#== normal Func ==#

def multiTochr(data) :
	ret_data = ""
	for i in range(0,len(data),2) :
		ret_data += chr(ord(data[i:i+1]))
	return ret_data


#======================================#

#== get a Sector Allocation Table ==#

SATable = getSAT(totalData,sHeader)
#print(SATable)###

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

streamData = []

for i in range(len(aDirEntry)) :
	hwFile.write("DirName : "+aDirEntry[i].DirName+"\n")
	if aDirEntry[i].StreamSize >= MINSECSIZE or aDirEntry[i].Type == 5:
		dirData, dirSecCount = followSecs(SATable,aDirEntry[i].SecID_Start,totalData)
	else:
		dirData, dirSecCount = followSSecs(SSATable,aDirEntry[i].SecID_Start,SSecData)

	streamData.append(dirData)	
	hwFile.write(dirData+"\n\n")

hwFile.close()	

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
		Table2Num = i	

#########################

ntFibBase = namedtuple('FibBase',ruleName_FibBase)
sFibBase = ntFibBase._make(unpack(ruleChar_FibBase,streamData[wordStreamNum][:32]))

ntFibBase_opt = namedtuple('FibBaseOpt',ruleName_FibBaseOpt)
sFibBase_opt = ntFibBase_opt._make(bitParse(ruleChar_FibBaseOpt,sFibBase.opt))

ntFibRgFcLcb97 = namedtuple('FibRgFcLcb97',ruleName_FibRgFcLcb97)
sFibRgFcLcb97 = ntFibRgFcLcb97._make(unpack(ruleChar_FibRgFcLcb97,streamData[wordStreamNum][0x9A:0x9A+0x2e8]))

if sFibBase_opt.fWhichTblStm == 1 :
	tableNum = Table1Num
else:
	tableNum = Table0Num
TableStream = streamData[tableNum]

dataPlcBteChpx = TableStream[sFibRgFcLcb97.fcPlcfBteChpx:sFibRgFcLcb97.fcPlcfBteChpx+sFibRgFcLcb97.lcbPlcfBteChpx]
ntPlcBteChpx = namedtuple('PlcBteChpx',ruleName_PlcBteChpx)
sPlcBteChpx = ntPlcBteChpx._make(unpack(ruleChar_PlcBteChpx,dataPlcBteChpx))

ntPnFkpChpx = namedtuple('PnFkpChpx',ruleName_PnFkpChpx)
sPnFkpChpx = ntPnFkpChpx._make(bitParse(ruleChar_PnFkpChpx,sPlcBteChpx.aPnBteChpx))
pStart_ChpxFkp = sPnFkpChpx.pn*512

dataChpxFkp = streamData[wordStreamNum][pStart_ChpxFkp:pStart_ChpxFkp+512]
ntChpxFkp = namedtuple('ChpxFkp',ruleName_ChpxFkp)
sChpxFkp = ntChpxFkp._make(makeChpxFkp(dataChpxFkp,sPlcBteChpx.aFC1))

#======================================#

#== get the information of Chpx structure ==#

dataRgfc, dataRgb = getDataChpxFkp(streamData[wordStreamNum],sChpxFkp,pStart_ChpxFkp)

#print(dataRgb)
for i in range(len(dataRgb)) :
	Pri,sprm = makePri(dataRgb[i][1])
	print("# [%d] #==========================================#"%i)
	print(Pri)
	

#======================================#

#print sFibBase
#print sFibBase_opt
#print sFibRgFcLcb97
#print sPlcBteChpx
#print sPnFkpChpx
#print sChpxFkp
#print dataRgfc
#print dataRgb
#"""
count = 0
for i in dataRgfc :
	print("[%d]"%count+" : "+i)
	count += 1
#"""

print "fEncrypted : %d"%sFibBase_opt.fEncrypted
print "fObfuscated : %d"%sFibBase_opt.fObfuscated

"""
sprmCPicLocation(0x6A03) - OLE 스토리지의 이름을 가리킨다.
sprmCFOle2(0x080A)
sprmCRsidText(0x6816)
sprmCHps(0x4A43)
sprmCHpsKern(0x484B)
sprmCFSpec(0x0855)
sprmCFObj(0x0856)
sprmCHpsBi(0x4A61)
sprmCIdctHint(0x286F)
"""