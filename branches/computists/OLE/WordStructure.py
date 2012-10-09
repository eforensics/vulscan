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

def makeChpxFkp(data,end):
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

def getDataChpxFkp(stream,sChpxFkp,pStart_ChpxFkp):
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
	spraTable = {0:1, 1:1, 2:2, 3:4, 4:2, 5:2, 6:0, 7:3} ## Exactly, the first one, the zero, has 1 bit toggle operand. not 1 byte operand
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
		print(operandCount)
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

ruleChar_FIB = '32sH28sH88sH?H?'
ruleName_FIB = ('base csw fibRgW cslw fibRgLw cbRgFcLcb fibRgFcLcbBlob cswNew fibRgCswNew')

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

ruleChar_FibRgFcLcb97 = '186L'
ruleName_FibRgFcLcb97 = ('fcStshfOrig lcbStshfOrig fcStshf lcbStshf fcPlcffndRef lcbPlcffndRef fcPlcffndTxt lcbPlcffndTxt fcPlcfandRef lcbPlcfandRef fcPlcfandTxt lcbPlcfandTxt fcPlcfSed lcbPlcfSed fcPlcPad lcbPlcPad fcPlcfPhe lcbPlcfPhe fcSttbfGlsy lcbSttbfGlsy fcPlcfGlsy lcbPlcfGlsy fcPlcfHdd lcbPlcfHdd fcPlcfBteChpx lcbPlcfBteChpx fcPlcfBtePapx lcbPlcfBtePapx fcPlcfSea lcbPlcfSea fcSttbfFfn lcbSttbfFfn fcPlcfFldMom lcbPlcfFldMom fcPlcfFldHdr lcbPlcfFldHdr fcPlcfFldFtn lcbPlcfFldFtn fcPlcfFldAtn lcbPlcfFldAtn fcPlcfFldMcr lcbPlcfFldMcr fcSttbfBkmk lcbSttbfBkmk fcPlcfBkf lcbPlcfBkf fcPlcfBkl lcbPlcfBkl fcCmds lcbCmds fcUnused1 lcbUnused1 fcSttbfMcr lcbSttbfMcr fcPrDrvr lcbPrDrvr fcPrEnvPort lcbPrEnvPort fcPrEnvLand lcbPrEnvLand fcWss lcbWss fcDop lcbDop fcSttbfAssoc lcbSttbfAssoc fcClx lcbClx fcPlcfPgdFtn lcbPlcfPgdFtn fcAutosaveSource lcbAutosaveSource fcGrpXstAtnOwners lcbGrpXstAtnOwners fcSttbfAtnBkmk lcbSttbfAtnBkmk fcUnused2 lcbUnused2 fcUnused3 lcbUnused3 fcPlcSpaMom lcbPlcSpaMom fcPlcSpaHdr lcbPlcSpaHdr fcPlcfAtnBkf lcbPlcfAtnBkf fcPlcfAtnBkl lcbPlcfAtnBkl fcPms lcbPms fcFormFldSttbs lcbFormFldSttbs fcPlcfendRef lcbPlcfendRef fcPlcfendTxt lcbPlcfendTxt fcPlcfFldEdn lcbPlcfFldEdn fcUnused4 lcbUnused4 fcDggInfo lcbDggInfo fcSttbfRMark lcbSttbfRMark fcSttbfCaption lcbSttbfCaption fcSttbfAutoCaption lcbSttbfAutoCaption fcPlcfWkb lcbPlcfWkb fcPlcfSpl lcbPlcfSpl fcPlcftxbxTxt lcbPlcftxbxTxt fcPlcfFldTxbx lcbPlcfFldTxbx fcPlcfHdrtxbxTxt lcbPlcfHdrtxbxTxt fcPlcffldHdrTxbx lcbPlcffldHdrTxbx fcStwUser lcbStwUser fcSttbTtmbd lcbSttbTtmbd fcCookieData lcbCookieData fcPgdMotherOldOld lcbPgdMotherOldOld fcBkdMotherOldOld lcbBkdMotherOldOld fcPgdFtnOldOld lcbPgdFtnOldOld fcBkdFtnOldOld lcbBkdFtnOldOld fcPgdEdnOldOld lcbPgdEdnOldOld fcBkdEdnOldOld lcbBkdEdnOldOld fcSttbfIntlFld lcbSttbfIntlFld fcRouteSlip lcbRouteSlip fcSttbSavedBy lcbSttbSavedBy fcSttbFnm lcbSttbFnm fcPlfLst lcbPlfLst fcPlfLfo lcbPlfLfo fcPlcfTxbxBkd lcbPlcfTxbxBkd fcPlcfTxbxHdrBkd lcbPlcfTxbxHdrBkd fcDocUndoWord9 lcbDocUndoWord9 fcRgbUse lcbRgbUse fcUsp lcbUsp fcUskf lcbUskf fcPlcupcRgbUse lcbPlcupcRgbUse fcPlcupcUsp lcbPlcupcUsp fcSttbGlsyStyle lcbSttbGlsyStyle fcPlgosl lcbPlgosl fcPlcocx lcbPlcocx fcPlcfBteLvc lcbPlcfBteLvc dwLowDateTime dwHighDateTime fcPlcfLvcPre10 lcbPlcfLvcPre10 fcPlcfAsumy lcbPlcfAsumy fcPlcfGram lcbPlcfGram fcSttbListNames lcbSttbListNames fcSttbfUssr lcbSttbfUssr')

#======================================#


#== PlcBteChpx Structure ==#
# structure size : 0xc(variable)

ruleChar_PlcBteChpx = 'LLL'
ruleName_PlcBteChpx = ('aFC0 aFC1 aPnBteChpx')

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
	