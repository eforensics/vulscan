# External Module Import
import re, os
from struct import unpack
from collections import namedtuple
from traceback import format_exc

# Internal Module Import
from Common import *
	
def fnIsOLE(sData) :
	# [IN]		sData
	# [OUT]		"OLE" / False
	sbuffer = ""
	try :
		if os.path.exists(sData) :
			sbuffer = CFile.fnReadFile(sData)[:200]
	except :
		sbuffer = sData[:200]
			
	if ((sbuffer[0:4] == "\xd0\xcf\x11\xe0") and (sbuffer[4:8] == "\xa1\xb1\x1a\xe1")) or (sbuffer[0:4] == "\xd0\xcd\x11\xe0") : 
		return "OLE"
	return False
def fnMap(sName, sbuffer, sRuleName, sRulePattern) :
	tName = namedtuple(sName, sRuleName)
	return tName._make( unpack(sRulePattern, sbuffer) )
def fnRead(sbuffer, nSecID) :
	nSecPos = (nSecID + 1) * SIZE_SECTOR
	return sbuffer[nSecPos:nSecPos + SIZE_SECTOR]
	
class CSECTOR():
	@classmethod
	def fnReadMSAT(cls, sbuffer):
		nPos_MSAT_nSecID = 0x44	# tHeader.MSATSecID
		nPos_MSAT_Num = 0x48	# tHeader.NumMSAT
		nPos_MSAT = 0x4C		# tHeader.MSAT

		lPos_MSAT = []
		sMSAT = sbuffer[nPos_MSAT:SIZE_SECTOR]
		
		nMSATNum = CBuffer.fnReadData(sbuffer, nPos_MSAT_Num, 4)
		while nMSATNum > 0 :
			sMSAT += fnRead(sbuffer, nPos_MSAT_nSecID)
			nPos_MSAT_nSecID = CUtil.fnReadData(sMSAT, len(sMSAT) -4, 4)
			if nPos_MSAT_nSecID > 0 :
				sMSAT = sMSAT[:-4]
			nMSATNum -= 1
			
		return CUtil.fnPack(sMSAT, 4)
	@classmethod
	def fnReadSAT(cls, sbuffer, tMSAT):
		sSAT = ""
		for nSecID in tMSAT :
			if nSecID < 0 :
				break
			sSAT += fnRead(sbuffer, nSecID)
		
		return CUtil.fnPack(sSAT, 4)	
	@classmethod
	def fnReadSSAT(cls, sbuffer, SAT, nSecID):
		sSector = cls.fnReadSector(sbuffer, SAT, nSecID)
		if sSector == False :
			return False
		
		return CUtil.fnPack(sSector, 4)
	@classmethod
	def fnReadSector(cls, sbuffer, tTable, nSecID):
		lSecID = cls.fnReadSecIDListByTable(tTable, nSecID)
		if lSecID == [] :
			return False
		
		return cls.fnReadByList(sbuffer, lSecID)
	@classmethod
	def fnReadByList(cls, sbuffer, lSecID):	
		sSector = ""
		for nSecID in lSecID :
			sSector += fnRead(sbuffer, nSecID)
		
		return sSector			
	@classmethod
	def fnReadSecIDListByTable(cls, tTable, nSecID):
		lSecID = []
		while nSecID >= 0 :
			lSecID.append( nSecID )
			nSecID = tTable[nSecID]
	
		return lSecID
	@classmethod
	def fnReadByDirectory(cls, sbuffer, lDirectory, tSAT, tSSAT, sName):
		for tDirectory in lDirectory :
			if CUtil.fnExtractAlphaNumber(tDirectory.DirName).find(sName) == -1 :
				continue
			
			if tDirectory.Type == 5 or tDirectory.StreamSize >= 0x1000 :
				tTable = tSAT
			else :
				tTable = tSSAT
							
			return cls.fnReadSector(sbuffer, tTable, tDirectory.SecID_Start)
		return False
	
class CPRINTOLE():
	@classmethod
	def fnHeader(cls, tHeader) :
		pass
	@classmethod
	def fnDirectory(cls, tDirectory) :
		if tDirectory.DirName == "" :
			return True
	
		print "\t" + "=" * 130
		print "\t" + "%-33s%-11s%-11s%-11s%-11s%-11s%-11s%-11s%-10s%-8s" % ("DirectoryName", "szName", "Type", "Color", "Left", "Right", "Root", "Flags", "SecID", "szStream")
		# print "\t" + "-" * 130
		
		for nIndex in range( len(tDirectory) ) :
			if nIndex == 0 :
				print "\t" + "%-30s" % CUtil.fnExtractAlphaNumber(tDirectory[0]),
			elif nIndex == 7 or nIndex == 9 or nIndex == 10 or nIndex ==13 :
				continue
			else :
				print "0x%08X" % tDirectory[nIndex],
		print ""
		
		print "\t" + "=" * 130
		return True
		
SIZE_SECTOR			= 0x200
SIZE_SECTOR_SHORT	= 0x40
		
SIZE_OLE_HEADER		= 0x200
SIZE_OLE_DIRECTORY	= 0x80

RULE_OLE_HEADER_NAME	= ('Signature UID Revision Version ByteOrder szSector szShort Reserved1 NumSAT DirSecID Reserved2 szMinSector SSATSecID NumSSAT MSATSecID NumMSAT MSAT')
RULE_OLE_HEADER_PATTERN	= '=8s16s5h10s8l436s'

RULE_OLE_DIRECTORY_NAME	= ('DirName Size Type Color LeftChild RightChild RootChild UID Flags createTime LastModifyTime SecID_Start StreamSize Reserved')
RULE_OLE_DIRECTORY_PATTERN	= '=64sH2B3l16sl8s8s3l'

RULE_HWP_HEADER_NAME = ('Signature Version Property Reserved')
RULE_HWP_HEADER_PATTERN = '=32s2l216s'


