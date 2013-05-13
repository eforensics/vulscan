# External Module Import
from traceback import format_exc

# Internal Module Import
from Common import *

class CDOC():
	@classmethod
	def fnCheckDOC(cls, sWordDocumet):
		nPos = SIZE_FIBBASE
		nCSW = CBuffer.fnReadData(sWordDocument, nPos, 2)
		if nCSW != 0x000E :
			print "[Failure] fnCheckDOC( nCSW : 0x%02X )" % nCSW
			return False
			
		nPos += 2 + SIZE_FIBW97
		nCSLW = CBuffer.fnReadData(sWordDocument, nPos, 2)
		if nCSLW != 0x0016 :
			print "[Failure] fnCheckDOC( nCSLW : 0x%02X )" % nCSLW
			return False
			
		nPos += 2 + SIZE_FIBLW97
		ncbRgFcLcb = CBuffer.fnReadData(sWordDocument, nPos, 2)
		if not ncbRgFcLcb in [0x005D, 0x006C, 0x0088, 0x00A4, 0x00B7] :
			print "[Failure] fnCheckDOC( ncbRgFcLcb : 0x%02X)" % ncbRgFcLcb
			return False
			
		nPos += 2 + ncbRgFcLcb * 8
		ncswNew = CBuffer.fnReadData(sWordDocument, nPos, 2)
		if not ncswNew in [0x0000, 0x0002, 0x0005] :
			print "[Failure] fnCheckDOC( ncswNew : 0x%02X )" % ncswNew
			return False
		
		return True
	@classmethod
	def fnFib(cls, sWordDocument):
		nPos_cswNew = 0x5ba
		nPos_FibBase = 0x2
		
		ncswNew = CBuffer.fnReadData(sWordDocument, nPos_cswNew, 2)
		nFib = CBuffer.fnReadData(sWordDocument, nPos_FibBase, 4)
		
		if ncswNew == 0x0000 :
			return nFib, 0
		else :
			return CBuffer.fnReadData(sWordDocument, nPos_cswNew + 2, 2), 2
	@classmethod
	def fnTableName(cls, sWordDocument):
		nPos_Flag = 0xA
		nTable = CUtil.fnBitParse( CBuffer.fnReadData(sWordDocument, nPos_Flag, 2), 9, 1 )
		return "%sTable" % nTable
		
	@classmethod
	def fnFibBase(cls, sWordDocument):
		return fnMap("FibBase", sWordDocument[nPos:nPos + SIZE_FIBBASE], RULE_OFFICE_FIBBASE_NAME, RULE_OFFICE_FIBBASE_PATTERN)
	@classmethod
	def fnFibRgW97(cls, sWordDocument):
		nPos = SIZE_FIBBASE + 2
		return fnMap("FibRgW97", sWordDocument[nPos:nPos + SIZE_FIBW97], RULE_OFFICE_FIBW97_NAME, RULE_OFFICE_FIBW97_PATTERN)
	@classmethod
	def fnFibRgLw97(cls, sWordDocument):
		nPos = SIZE_FIBBASE + 2 + SIZE_FIBW97 + 2
		return fnMap("FibRgLw97", sWordDocument[nPos:nPos + SIZE_FIBLW97], RULE_OFFICE_FIBLW97_NAME, RULE_OFFICE_FIBLW97_PATTERN)
		
class CPRINTDOC():
	@classmethod
	def fnFibBase(cls, sWordDocument):
		pass
	@classmethod
	def fnFibRgW97(cls, sWordDocument):
		pass
	@classmethod
	def fnFibRgLw97(cls, sWordDocument):
		pass	
	@classmethod
	def fnFibRgFcLcb(cls, tFibRgFcLcb):
		pass

		
SIZE_FIBBASE = 0x20		
RULE_OFFICE_FIBBASE_NAME = ('wIdent nFib unused lid pnNext Flag nFibBack lKey envr envtFlag reserved3 reserved4 reserved5 reserved6')
RULE_OFFICE_FIBBASE_PATTERN = '=7H1L2B2H2L'

SIZE_FIBW97 = 0x1c
RULE_OFFICE_FIBW97_NAME = ('reserved1 reserved2 reserved3 reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 lidFE')
RULE_OFFICE_FIBW97_PATTERN = '=14H'

SIZE_FIBLW97 = 0x58
RULE_OFFICE_FIBLW97_NAME = ('cbMac reserved1 reserved2 ccpText ccpFtn ccpHdd reserved3 ccpAtn ccpEdn ccpTxbx ccpHdrTxbx reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 reserved14')
RULE_OFFICE_FIBLW97_PATTERN = '=22L'

lFib = [0x00C1, 0x00D9, 0x0101, 0x010C, 0x0112]
lVer = [97, 2000, 2002, 2003, 2007]

SIZE_FIBRGFCLCB_97 = 0x2E8
RULE_OFFICE_FIBRGFCLCB_NAME_97 = ('fcStshfOrig lcbStshfOrig fcStshf lcbStshf fcPlcffndRef lcbPlcffndRef fcPlcffndTxt lcbPlcffndTxt fcPlcfandRef lcbPlcfandRef fcPlcfandTxt lcbPlcfandTxt fcPlcfSed lcbPlcfSed fcPlcPad lcbPlcPad fcPlcfPhe lcbPlcfPhe fcSttbfGlsy lcbSttbfGlsy fcPlcfGlsy lcbPlcfGlsy fcPlcfHdd lcbPlcfHdd fcPlcBteChpx lcbPlcBteChpx fcPlcfBtePapx lcbPlcfBtePapx fcPlcfSea lcbPlcfSea fcSttbfFfn lcbSttbfFfn fcPlcfFldMom lcbPlcfFldMom fcPlcfFldHdr lcbPlcfFldHdr fcPlcfFldFtn lcbPlcfFldFtn fcPlcfFldAtn lcbPlcfFldAtn fcPlcfFldMcr lcbPlcfFldMcr fcSttbfBkmk lcbSttbfBkmk fcPlcfBkf lcbPlcfBkf fcPlcfBkl lcbPlcfBkl fcCmds lcbCmds fcUnused1 lcbUnused1 fcSttbfMcr lcbSttbfMcr fcPrDrvr lcbPDrvr fcPrEnvPort lcbPrEnvPort fcPrEnvLand lcbPrEnvLand fcWss lcbWss fcDop lcbDop fcSttbfAssoc lcbSttbfAssoc fcClx lcbClx fcPlcfPgdFtn lcbPlcfPgdFtn fcAutosaveSource lcbAutosaveSource fcGrpXstAtnOwners lcbGrpXstAtnOwners fcSttbfAtnBkmk lcbSttbfAtnBkmk fcUnused3 lcbUnused3 fcUnused2 lcbUnused2 fcPlcSpaMon lcbPlcSpaMon fcPlcSpaHdr lcbPlcSpaHdr fcPlcAtnBkf lcbPlcAtnBkf fcPlcAtnBkl lcbPlcAtnBkl fcPms lcbPms fcFormFldSttbs lcbFormFldSttbs fcPlcfendRef lcbPlcfendRef fcPlcfendTxt lcbPlcfendTxt fcPlcfFldEdn lcbPlcfFldEdn fcUnused4 lcbUnused4 fcDggInfo lcbDggInfo fcSttbfRMark lcbSttbfRMark fcSttbfCaption lcbSttbfCaption fcSttbfAutoCaption lcbSttbfAutoCaption fcPlcWkb lcbPlcWkb fcPlcfSpl lcbPlcfSpl fcPlcftxbxTxt lcbPlcftxbxTxt fcPlcfFldTxbx lcbPlcfFldTxbx fcPlcfHdrtxbxTxt lcbPlcfHdrtxbxTxt fcPlcffldHdrTxdx lcbPlcffldHdrTxdx fcStwUser lcbStwUser fcSttbTtmbd lcbSttbTtmbd fcCookieData lcbCookieData fcPgdMotherOldOld lcbPgdMotherOldOld fcBkdMotherOldOld lcbBkdMotherOldOld fcPgdFtnOldOld lcbPgdFtnOldOld fcBkdFtnOldOld lcbBkdFtnOldOld fcPgdEdnOldOld lcbPgdEdnOldOld fcBkdEdnOldOld lcbBkdEdnOldOld fcSttbfIntlFld lcbSttbfIntlFld fcRouteSlip lcbRouteSlip fcSttbSavedBy lcbSttbSavedBy fcSttbFnm lcbSttbFnm fcPlfLst lcbPlfLst fcPlfLfo lcbPlfLfo fcPlcTxbxBkd lcbPlcTxbxBkd fcPlcfTxbxHdrBkd lcbPlcfTxbxHdrBkd fcDocUndoWord9 lcbDocUndoWord9 fcRgbUse lcbRgbUse fcUsp lcbUsp fcUskf lcbUskf fcPlcupcRgbUse lcbPlcupcRgbUse fcPlcupcUsp lcbPlcupcUsp fcSttbGlsyStyle lcbSttbGlsyStyle fcPlgosl lcbPlgosl fcPlcocx lcbPlcocx fcPlcfBteLvc lcbPlcfBteLvc dwLowDateTime dwHighDateTime fcPlcfLvcPe10 lcbPlcfLvcPe10 fcPlcfAsumy lcbPlcfAsumy fcPlcfGram lcbPlcfGram fcSttbListNames lcbSttbListNames fcSttbfUssr lcbSttbfUssr')
RULE_OFFICE_FIBRGFCLCB_PATTERN_97 = '=186L'

SIZE_FIBRGFCLCB_2000 = 0x78
RULE_OFFICE_FIBRGFCLCB_NAME_2000 = ('fcPlcfTch lcbPlcfTch fcRmdThreading lcbRmdThreading fcMid lcbMid fcSttbRgtplc lcbSttbRgtplc fcMsoEnvelope lcbMsoEnvelope fcPlcfLad lcbPlcfLad fcRgDofr lcbRgDofr fcPlcosl lcbPlcosl fcPlcfCookieOld lcbPlcfCookieOld fcPgdMotherOld lcbPgdMotherOld fcBkdMotherOld lcbBkdMotherOld fcPgdFtnOld lcbPgdFtnOld fcBkdFtnOld lcbBkdFtnOld fcPgdEdnOld lcbPgdEdnOld fcBkdEdnOld lcbBkdEdnOld')
RULE_OFFICE_FIBRGFCLCB_PATTERN_2000 = '=30L'

SIZE_FIBRGFCLCB_2002 = 0xE0
RULE_OFFICE_FIBRGFCLCB_NAME_2002 = ('fcUnused1 lcbUnused1 fcPlcfPgp lcbPlcfPgp fcPlcfuim lcbPlcfuim fcPlfguidUim lcbPlfguidUim fcAtrdExtra lcbAtrdExtra fcPlrsid lcbPlrsid fcSttbfBkmkFactoid lcbSttbfBkmkFactoid fcPlcfBkfFactoid lcbPlcfBkfFactoid fcPlcfcookie lcbPlcfcookie fcPlcfBklFactoid lcbPlcfBklFactoid fcFactoidData lcbFactoidData fcDocUndo lcbDocUndo fcSttbfBkmkFcc lcbSttbfBkmkFcc fcPlcfBkfFcc lcbPlcfBkfFcc fcPlcfBklFcc lcbPlcfBklFcc fcSttbfbkmkBPRepairs lcbSttbfbkmkBPRepairs fcPlcfbkfBPRepairs lcbPlcfbkfBPRepairs fcPlcfbklBPRepairs lcbPlcfbklBPRepairs fcPmsNew lcbPmsNew fcODSO lcbODSO fcPlcfpmiOldXP lcbPlcfpmiOldXP fcPlcfpmiNewXP lcbPlcfpmiNewXP fcPlcfpmiMixedXP lcbPlcfpmiMixedXP fcUnused2 lcbUnused2 fcPlcffactoid lcbPlcffactoid fcPlcflvcOldXP lcbPlcflvcOldXP fcPlcflvcNewXP lcbPlcflvcNewXP fcPlcflvcMixedXP lcbPlcflvcMixedXP')
RULE_OFFICE_FIBRGFCLCB_PATTERN_2002 = '=56L'

SIZE_FIBRGFCLCB_2003 = 0xE0
RULE_OFFICE_FIBRGFCLCB_NAME_2003 = ('fcHplxsdr lcbHplxsdr fcSttbfBkmkSdt lcbSttbfBkmkSdt fcPlcfBkfSdt lcbPlcfBkfSdt fcPlcfBklSdt lcbPlcfBklSdt fcCustomXForm lcbCustomXForm fcSttbfBkmkProt lcbSttbfBkmkProt fcPlcfBkfProt lcbPlcfBkfProt fcPlcfBklProt lcbPlcfBklProt fcSttbProtUser lcbSttbProtUser fcUnused lcbUnused fcPlcfpmiOld lcbPlcfpmiOld fcPlcfpmiOldInline lcbPlcfpmiOldInline fcPlcfpmiNew lcbPlcfpmiNew fcPlcfpmiNewInline lcbPlcfpmiNewInline fcPlcflvcOld lcbPlcflvcOld fcPlcflvcOldInline lcbPlcflvcOldInline fcPlcflvcNew lcbPlcflvcNew fcPlcflcxNewInline lcbPlcflcxNewInline fcPgdMother lcbPgdMother fcBkdMother lcbBkdMother fcAfdMother lcbAfdMother fcPgdFtn lcbPgdFtn fcBkdFtn lcbBkdFtn fcAfdFtn lcbAfdFtn fcPgdEdn lcbPgdEdn fcBkdEdn lcbBkdEdn fcAfdEdn lcbAfdEdn fcAfd lcbAfd')
RULE_OFFICE_FIBRGFCLCB_PATTERN_2003 = '=56L'

SIZE_FIBRGFCLCB_2007 = 0x98
RULE_OFFICE_FIBRGFCLCB_NAME_2007 = ('fcPlcfmthd lcbPlcfmthd fcSttbfBkmkMoveFrom lcbSttbfBkmkMoveFrom fcPlcfBkfMoveFrom lcbPlcfBkfMoveFrom fcPlcfBklMoveFrom lcbPlcfBklMoveFrom fcSttbfBkmkMoveTo lcbSttbfBkmkMoveTo fcPlcfBkfMoveTo lcbPlcfBkfMoveTo fcPlcfBklMoveTo lcbPlcfBklMoveTo fcUnused1 lcbUnused1 fcUnused2 lcbUnused2 fcUnused3 lcbUnused3 fcSttbfBkmkArto lcbSttbfBkmkArto fcPlcfBkfArto lcbPlcfBkfArto fcPlcfBklArto lcbPlcfBklArto fcArtoData lcbArtoData fcUnused4 lcbUnused4 fcUnused5 lcbUnused5 fcUnused6 lcbUnused6 fcOssTheme lcbOssTheme fcColorSchemeMapping lcbColorSchemeMapping')
RULE_OFFICE_FIBRGFCLCB_PATTERN_2007 = '=38L'

