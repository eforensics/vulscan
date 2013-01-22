# -*- coding: utf-8 -*-
from struct import *
from collections import namedtuple

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

def makeChpxFkp(data):
	dCount = 0
	retData = []
	rgfc = []
	rgb = []
	pStart = 0
	pEnd = 0
	crun = unpack('<B',data[511])[0]
	if crun < 0x1 or crun > 0x65 :
		print("crun value in ChpxFkp structure has wrong")

	for i in range(crun+1) :
		pStart = pEnd
		pEnd += 4
		dCount += 1

		value = unpack('<L',data[pStart:pEnd])[0]
		rgfc.append(value)

	rgbCount = dCount - 1
	while(True) :
		value = unpack('<B',data[pEnd])[0]
		if rgbCount == 0 :
			break
		else:
			pEnd += 1
			rgbCount += -1
			rgb.append(value)

	retData.append(rgfc)
	retData.append(rgb)
	retData.append(crun)
	
	return retData

def getDataChpx(stream,sChpxFkp,pStart_ChpxFkp):
	dataRgfc = []
	dataRgb = []
	for i in range(sChpxFkp.crun) :
		dataRgfc.append(stream[sChpxFkp.rgfc[i]:sChpxFkp.rgfc[i+1]])

	for i in range(sChpxFkp.crun) :
		pChpx = pStart_ChpxFkp+sChpxFkp.rgb[i]*2
		cb = unpack('<B',stream[pChpx])[0]
		pGrpprl = pChpx+1
		grpprl = stream[pGrpprl:pGrpprl+cb]

		dataRgb.append([cb,grpprl])

	return dataRgfc, dataRgb


def checkSpra (data):
	spraTable = {0:1, 1:1, 2:2, 3:4, 4:2, 5:2, 6:0, 7:3} ## the first one, the zero, has 1 bit toggle operand. not a byte operand
	for i in spraTable :
		if data == i:
#			print("spra : %d"%i)
			return spraTable[i]	

def makePri (data):
	sprmSize = 2
	Pri = []
	pSprm = 0
	while(True) :
#		print(pSprm)
		dataSprm = unpack('<H',data[pSprm:pSprm+sprmSize])[0]
		sprm = bitParse(ruleChar_Sprm,dataSprm)
		operandCount = checkSpra(sprm[3])
#		print("operandCount : %d"%operandCount)
		offset = pSprm + sprmSize + operandCount
		operand = data[pSprm+sprmSize:offset]
		Pri.append([dataSprm, operand, sprm])
		if offset == len(data) :
			break
		else :
			pSprm = offset


	return Pri, sprm
		
#======================================#

#== FIB Structure ==#
# offset : 0x00 in WordDocument Stream
# replace func getFIB() in main with this rules.

#ruleChar_FIB = '32sH28sH88sH?H?'
#ruleName_FIB = ('base csw fibRgW cslw fibRgLw cbRgFcLcb FibRgFcLcbBlob cswNew fibRgCswNew')

def getFIB (data):
	size_Base = 32
	offset_Base = 0
	pStart = offset_Base
	pEnd = size_Base
	
	mValue = [2,4,8,2]

	pointer = [0,32]
	pOffset = [32]

	for i in range(0,4) : ###
		pStart = pEnd
		pEnd += 2
		pointer.append(pEnd)
		offset = unpack('<H',data[pStart:pEnd])[0] * mValue[i]
		pEnd = pEnd + offset
		pointer.append(pEnd)
		pOffset.append(offset)
		
	ruleChar_FIB = '32sH28sH88sH' + str(pOffset[3]) + 'sH' + str(pOffset[4]) + 's' ### variable 한 FibRgFcLcbBlob 때문에 flexible 한 데이터를 파싱하기 위해서 fix 할 수 없다.
	ruleName_FIB = ('base csw fibRgW cslw fibRgLw cbRgFcLcb FibRgFcLcbBlob cswNew fibRgCswNew')

	ntFIB = namedtuple('FIB',ruleName_FIB)
	sFIB = ntFIB._make(unpack(ruleChar_FIB,data[:pointer[9]]))

	return sFIB

#======================================#


#== FIB Base Structure ==#
# offset : 0x00 at FIB
# structure size : 0x2e8

ruleChar_FibBase = '<2s6HL3H2L'
ruleName_FibBase = ('wIdent nFib unused lid pnNext opt nFibBack lKey opt2 reserved3 reserved4 reserved5 reserved6')

ruleChar_FibBaseOpt = "1,1,1,1,4,1,1,1,1,1,1,1,1"
ruleName_FibBaseOpt = ('fDot fGlsy fComplex fHasPic cQuickSaves fEncrypted fWhichTblStm fReadOnlyRecommended fWriteReservation fExtChar fLoadOverride fFarEast fObfuscated')

#======================================#


#== FibRgFcLcb97 Structure ==#
# offset : 0x9A
# structure size : 0x2e8

def FibRgFcLcb97 (data):
	ruleChar_FibRgFcLcb97 = '186L'
	ruleName_FibRgFcLcb97 = ('fcStshfOrig lcbStshfOrig fcStshf lcbStshf fcPlcffndRef lcbPlcffndRef fcPlcffndTxt lcbPlcffndTxt fcPlcfandRef lcbPlcfandRef fcPlcfandTxt lcbPlcfandTxt fcPlcfSed lcbPlcfSed fcPlcPad lcbPlcPad fcPlcfPhe lcbPlcfPhe fcSttbfGlsy lcbSttbfGlsy fcPlcfGlsy lcbPlcfGlsy fcPlcfHdd lcbPlcfHdd fcPlcfBteChpx lcbPlcfBteChpx fcPlcfBtePapx lcbPlcfBtePapx fcPlcfSea lcbPlcfSea fcSttbfFfn lcbSttbfFfn fcPlcfFldMom lcbPlcfFldMom fcPlcfFldHdr lcbPlcfFldHdr fcPlcfFldFtn lcbPlcfFldFtn fcPlcfFldAtn lcbPlcfFldAtn fcPlcfFldMcr lcbPlcfFldMcr fcSttbfBkmk lcbSttbfBkmk fcPlcfBkf lcbPlcfBkf fcPlcfBkl lcbPlcfBkl fcCmds lcbCmds fcUnused1 lcbUnused1 fcSttbfMcr lcbSttbfMcr fcPrDrvr lcbPrDrvr fcPrEnvPort lcbPrEnvPort fcPrEnvLand lcbPrEnvLand fcWss lcbWss fcDop lcbDop fcSttbfAssoc lcbSttbfAssoc fcClx lcbClx fcPlcfPgdFtn lcbPlcfPgdFtn fcAutosaveSource lcbAutosaveSource fcGrpXstAtnOwners lcbGrpXstAtnOwners fcSttbfAtnBkmk lcbSttbfAtnBkmk fcUnused2 lcbUnused2 fcUnused3 lcbUnused3 fcPlcSpaMom lcbPlcSpaMom fcPlcSpaHdr lcbPlcSpaHdr fcPlcfAtnBkf lcbPlcfAtnBkf fcPlcfAtnBkl lcbPlcfAtnBkl fcPms lcbPms fcFormFldSttbs lcbFormFldSttbs fcPlcfendRef lcbPlcfendRef fcPlcfendTxt lcbPlcfendTxt fcPlcfFldEdn lcbPlcfFldEdn fcUnused4 lcbUnused4 fcDggInfo lcbDggInfo fcSttbfRMark lcbSttbfRMark fcSttbfCaption lcbSttbfCaption fcSttbfAutoCaption lcbSttbfAutoCaption fcPlcfWkb lcbPlcfWkb fcPlcfSpl lcbPlcfSpl fcPlcftxbxTxt lcbPlcftxbxTxt fcPlcfFldTxbx lcbPlcfFldTxbx fcPlcfHdrtxbxTxt lcbPlcfHdrtxbxTxt fcPlcffldHdrTxbx lcbPlcffldHdrTxbx fcStwUser lcbStwUser fcSttbTtmbd lcbSttbTtmbd fcCookieData lcbCookieData fcPgdMotherOldOld lcbPgdMotherOldOld fcBkdMotherOldOld lcbBkdMotherOldOld fcPgdFtnOldOld lcbPgdFtnOldOld fcBkdFtnOldOld lcbBkdFtnOldOld fcPgdEdnOldOld lcbPgdEdnOldOld fcBkdEdnOldOld lcbBkdEdnOldOld fcSttbfIntlFld lcbSttbfIntlFld fcRouteSlip lcbRouteSlip fcSttbSavedBy lcbSttbSavedBy fcSttbFnm lcbSttbFnm fcPlfLst lcbPlfLst fcPlfLfo lcbPlfLfo fcPlcfTxbxBkd lcbPlcfTxbxBkd fcPlcfTxbxHdrBkd lcbPlcfTxbxHdrBkd fcDocUndoWord9 lcbDocUndoWord9 fcRgbUse lcbRgbUse fcUsp lcbUsp fcUskf lcbUskf fcPlcupcRgbUse lcbPlcupcRgbUse fcPlcupcUsp lcbPlcupcUsp fcSttbGlsyStyle lcbSttbGlsyStyle fcPlgosl lcbPlgosl fcPlcocx lcbPlcocx fcPlcfBteLvc lcbPlcfBteLvc dwLowDateTime dwHighDateTime fcPlcfLvcPre10 lcbPlcfLvcPre10 fcPlcfAsumy lcbPlcfAsumy fcPlcfGram lcbPlcfGram fcSttbListNames lcbSttbListNames fcSttbfUssr lcbSttbfUssr')

	ntFibRgFcLcb97 = namedtuple('FibRgFcLcb97',ruleName_FibRgFcLcb97)
	sFibRgFcLcb97 = ntFibRgFcLcb97._make(unpack(ruleChar_FibRgFcLcb97,data))
	
	return sFibRgFcLcb97

#======================================#



#== FibRgFcLcb2000 Structure ==#
# offset : 
# structure size : 0x360(864)

def FibRgFcLcb2000 (data):
#	ruleChar_FibRgFcLcb2000 = '744s30L'
	ruleChar_FibRgFcLcb2000 = '30L'
	ruleName_FibRgFcLcb2000 = ('fcPlcfTch lcbPlcfTch fcRmdThreading lcbRmdThreading fcMid lcbMid fcSttbRgtplc lcbSttbRgtplc fcMsoEnvelope lcbMsoEnvelope fcPlcfLad lcbPlcfLad fcRgDofr lcbRgDofr fcPlcosl lcbPlcosl fcPlcfCookieOld lcbPlcfCookieOld fcPgdMotherOld lcbPgdMotherOld fcBkdMotherOld lcbBkdMotherOld fcPgdFtnOld lcbPgdFtnOld fcBkdFtnOld lcbBkdFtnOld fcPgdEdnOld lcbPgdEdnOld fcBkdEdnOld lcbBkdEdnOld')

	ntFibRgFcLcb2000 = namedtuple('FibRgFcLcb2000',ruleName_FibRgFcLcb2000)
	sFibRgFcLcb2000 = ntFibRgFcLcb2000._make(unpack(ruleChar_FibRgFcLcb2000,data))
	
	return sFibRgFcLcb2000

#======================================#


#== FibRgFcLcb2002 Structure ==#
# offset : 
# structure size : 0x440(1088)

def FibRgFcLcb2002 (data):
#	ruleChar_FibRgFcLcb2002 = '864s56L'
	ruleChar_FibRgFcLcb2002 = '56L'
	ruleName_FibRgFcLcb2002 = ('fcUnused1 lcbUnused1 fcPlcfPgp lcbPlcfPgp fcPlcfuim lcbPlcfuim fcPlfguidUim lcbPlfguidUim fcAtrdExtra lcbAtrdExtra fcPlrsid lcbPlrsid fcSttbfBkmkFactoid lcbSttbfBkmkFactoid fcPlcfBkfFactoid lcbPlcfBkfFactoid fcPlcfcookie lcbPlcfcookie fcPlcfBklFactoid lcbPlcfBklFactoid fcFactoidData lcbFactoidData fcDocUndo lcbDocUndo fcSttbfBkmkFcc lcbSttbfBkmkFcc fcPlcfBkfFcc lcbPlcfBkfFcc fcPlcfBklFcc lcbPlcfBklFcc fcSttbfbkmkBPRepairs lcbSttbfbkmkBPRepairs fcPlcfbkfBPRepairs lcbPlcfbkfBPRepairs fcPlcfbklBPRepairs lcbPlcfbklBPRepairs fcPmsNew lcbPmsNew fcODSO lcbODSO fcPlcfpmiOldXP lcbPlcfpmiOldXP fcPlcfpmiNewXP lcbPlcfpmiNewXP fcPlcfpmiMixedXP lcbPlcfpmiMixedXP fcUnused2 lcbUnused2 fcPlcffactoid lcbPlcffactoid fcPlcflvcOldXP lcbPlcflvcOldXP fcPlcflvcNewXP lcbPlcflvcNewXP fcPlcflvcMixedXP lcbPlcflvcMixedXP')

	ntFibRgFcLcb2002 = namedtuple('FibRgFcLcb2002',ruleName_FibRgFcLcb2002)
	sFibRgFcLcb2002 = ntFibRgFcLcb2002._make(unpack(ruleChar_FibRgFcLcb2002,data))
	
	return sFibRgFcLcb2002

#======================================#


#== FibRgFcLcb2003 Structure ==#
# offset : 
# structure size : 0x520(1312)

def FibRgFcLcb2003 (data):
#	ruleChar_FibRgFcLcb2003 = '1088s56L'
	ruleChar_FibRgFcLcb2003 = '56L'
	ruleName_FibRgFcLcb2003 = ('fcHplxsdr lcbHplxsdr fcSttbfBkmkSdt lcbSttbfBkmkSdt fcPlcfBkfSdt lcbPlcfBkfSdt fcPlcfBklSdt lcbPlcfBklSdt fcCustomXForm lcbCustomXForm fcSttbfBkmkProt lcbSttbfBkmkProt fcPlcfBkfProt lcbPlcfBkfProt fcPlcfBklProt lcbPlcfBklProt fcSttbProtUser lcbSttbProtUser fcUnused lcbUnused fcPlcfpmiOld lcbPlcfpmiOld fcPlcfpmiOldInline lcbPlcfpmiOldInline fcPlcfpmiNew lcbPlcfpmiNew fcPlcfpmiNewInline lcbPlcfpmiNewInline fcPlcflvcOld lcbPlcflvcOld fcPlcflvcOldInline lcbPlcflvcOldInline fcPlcflvcNew lcbPlcflvcNew fcPlcflvcNewInline lcbPlcflvcNewInline fcPgdMother lcbPgdMother fcBkdMother lcbBkdMother fcAfdMother lcbAfdMother fcPgdFtn lcbPgdFtn fcBkdFtn lcbBkdFtn fcAfdFtn lcbAfdFtn fcPgdEdn lcbPgdEdn fcBkdEdn lcbBkdEdn fcAfdEdn lcbAfdEdn fcAfd lcbAfd')

	ntFibRgFcLcb2003 = namedtuple('FibRgFcLcb2003',ruleName_FibRgFcLcb2003)
	sFibRgFcLcb2003 = ntFibRgFcLcb2003._make(unpack(ruleChar_FibRgFcLcb2003,data))
	
	return sFibRgFcLcb2003

#======================================#


#== FibRgFcLcb2007 Structure ==#
# offset : 
# structure size : 0x5b8(1464)

def FibRgFcLcb2007 (data):
#	ruleChar_FibRgFcLcb2007 = '1312s38L'
	ruleChar_FibRgFcLcb2007 = '38L'
	ruleName_FibRgFcLcb2007 = ('fcPlcfmthd lcbPlcfmthd fcSttbfBkmkMoveFrom lcbSttbfBkmkMoveFrom fcPlcfBkfMoveFrom lcbPlcfBkfMoveFrom fcPlcfBklMoveFrom lcbPlcfBklMoveFrom fcSttbfBkmkMoveTo lcbSttbfBkmkMoveTo fcPlcfBkfMoveTo lcbPlcfBkfMoveTo fcPlcfBklMoveTo lcbPlcfBklMoveTo fcUnused1 lcbUnused1 fcUnused2 lcbUnused2 fcUnused3 lcbUnused3 fcSttbfBkmkArto lcbSttbfBkmkArto fcPlcfBkfArto lcbPlcfBkfArto fcPlcfBklArto lcbPlcfBklArto fcArtoData lcbArtoData fcUnused4 lcbUnused4 fcUnused5 lcbUnused5 fcUnused6 lcbUnused6 fcOssTheme lcbOssTheme fcColorSchemeMapping lcbColorSchemeMapping')

	ntFibRgFcLcb2007 = namedtuple('FibRgFcLcb2007',ruleName_FibRgFcLcb2007)
	sFibRgFcLcb2007 = ntFibRgFcLcb2007._make(unpack(ruleChar_FibRgFcLcb2007,data))
	
	return sFibRgFcLcb2007

#======================================#
	

#== FibRgFcLcbBlob ==#

def getFibRgFcLcbBlob (cbRgFcLcb, data):

	FibRgFcLcbInfo = [[0x5d,FibRgFcLcb97,'FibRgFcLcb97'], [0x6c,FibRgFcLcb2000,'FibRgFcLcb2000'],[0x88,FibRgFcLcb2002,'FibRgFcLcb2002'],[0xa4,FibRgFcLcb2003,'FibRgFcLcb2003'],[0xb7,FibRgFcLcb2007,'FibRgFcLcb2007']] ### cbRgFcLcb | Func | Func name(str)
	pFirst = 0
	pEnd = 0	
	FibRgFcLcbData = []
	ruleName_FibRgFcLcb = ""

	for size,funcCall,funcName in FibRgFcLcbInfo :
		pFirst = pEnd
		pEnd = size * 8
		FibRgFcLcbData.append(funcCall(data[pFirst:pEnd]))
		ruleName_FibRgFcLcb += str(funcName) + " "
		if cbRgFcLcb == size:
			break
		else:
			continue

	ntFibRgFcLcbBlob = namedtuple('FibRgFcLcbBlob',(ruleName_FibRgFcLcb))
	sFibRgFcLcbBlob = ntFibRgFcLcbBlob._make(FibRgFcLcbData)

	return sFibRgFcLcbBlob

#======================================#

#== FibRgCswNew Structure ==#
# structure size : variable

def getFibRgCswNew (cswNew, data):
	if cswNew == 0 :	### cswNew 값이 0이면 FibRgCswNew 구조는 존재하지 않는다. 그러므로 None 을 리턴한다.
		return None
	else:	
		nFibNew = unpack('<H',data[:2])[0]
		if nFibNew != 0x112:
			size = 2
		else:
			size = 8	

		ruleChar_FibRgCswNew = '<H' + str(size) + 's'
		ruleName_FibRgCswNew = ('nFibNew rgCswNewData')

		ntFibRgCswNew = namedtuple('FibRgCswNew',ruleName_FibRgCswNew)
		sFibRgCswNew = ntFibRgCswNew._make(unpack(ruleChar_FibRgCswNew,data))

		return sFibRgCswNew
	
#======================================#

#== PlcBteChpx Structure ==#
# structure size : variable

def getPlcBteChpx (data):
	aFCCount = (len(data)/4)/2 + 1 ### aFC 와 aPnBteChpx 내의 배열의 크기는 명시되어있지 않다. ChpxFkp 와 같은 구조를 띄기 때문에 aFC가 aPnBteChpx 보다 한개 더 많다는 것으로 확인하자.
	aFCSize = aFCCount * 4
	aPnBteChpxCount = len(data)/4 - aFCCount
	aPnBteChpxSize = aPnBteChpxCount * 4


	aFC = [unpack('<L',data[i:i+4])[0] for i in range(0,aFCSize,4)]
	aPnBteChpx = [unpack('<L',data[i:i+4])[0] for i in range(aFCSize,aFCSize+aPnBteChpxSize,4)]

	ruleName_PlcBteChpx = ('aFC aPnBteChpx')

	ntPlcBteChpx = namedtuple('PlcBteChpx',ruleName_PlcBteChpx)
	sPlcBteChpx = ntPlcBteChpx._make([aFC,aPnBteChpx])

	return sPlcBteChpx

#======================================#


#== PnFkpChpx Structure ==#
# structure size : 0x4

ruleChar_PnFkpChpx = '22,10'
ruleName_PnFkpChpx = ('pn unused')

#======================================#


#== ChpxFkp Structure ==#
# structure size : 0x200

ruleChar_ChpxFkp = '^^;;;'
ruleName_ChpxFkp = ('rgfc rgb crun')

#======================================#


#== Chpx Structure ==#
# structure size : variable

ruleChar_Chpx = '8,?'
ruleName_Chpx = ('cb grpprl')

#======================================#


#== Prl Structure ==#
# structure size : variable

ruleChar_Prl = '16,?'
ruleName_Prl = ('sprm operand')

#======================================#


#== Sprm Structure ==#
# structure size : 0x2

ruleChar_Sprm = '9,1,3,3'
ruleName_Sprm = ('ispmd A sgc spra')

#======================================#
	
