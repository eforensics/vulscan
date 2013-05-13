# External Module Import
import os, sys
from optparse import OptionParser
from traceback import format_exc

# Internal Module Import
from VScan.parse.Common import *

try :
	from VScan.parse.PDF import *
	from VScan.parse.OLE import *
	from VScan.parse.RTF import fnIsRTF
	from VScan.parse.PE  import fnIsPE
	from VScan.parse.HWP import fnIsHWP
except :
	print "[Error] Internal Module Import"
	exit(0)
	
try :
	from VScan.parse.Office import *
except :
	print "[Error] Do not SCAN Office File"


def main():
	Init = CINIT()
	iOption = Init.fnGetOption()
	if iOption == False :
		return
	
	lFiles = Init.fnFindFile(iOption)
	if lFiles == False :
		return
	
	List = Init.fnSeperateFile(lFiles)
	
	Options = COPTIONS()
	if iOption.scan == True :
		Options.fnScan(iOption, List)
	
	if iOption.classify == True :
		Options.fnClassify(iOption, List)
	
	return
	
class CLIST():
	lFormat		= ["PDF", "RTF", "OLE", "HWP", "PE", "Unknown"]
	lKnown		= ["PDF", "RTF", "OLE", "HWP", "PE"]
	
	lPDF		= []
	lRTF		= []
	lOLE		= []
	lPE			= []
	lHWP		= []	# HWP 2.x/3.x
	lUnknown	= []
	
	lOLE_HWP	= []	# HWP 5.x
	lOLE_OFFICE	= []
	
	lExcept		= []
	
class CINIT():
	def __del__(self):
		return 
	def fnPrintLogo(self):
		print "-" * 65
		print "  VScan ver 1.0                   Copyright (c) 2013, Project1"
		print "-" * 65
		print ""
	def fnGetOption(self):
		self.fnPrintLogo()
		
		parser = OptionParser(usage="%prog -f <File> | -d <Dir> [Options]")
		parser.add_option("-f", "--file", help="File Path")
		parser.add_option("-d", "--dir", help="Directory Path")
		parser.add_option("-s", "--scan", action="store_true", help="Scan")
		parser.add_option("-c", "--classify", action="store_true", help="File Classifying")
		
		(iOption, args) = parser.parse_args()
		if len(sys.argv) < 2 :
			parser.print_help()
			return False
		return iOption	
	def fnFindFile(self, iOption):
		# [IN]		iOption
		# [OUT]
		#	- [SUCCESS]		lFiles
		#	- [FAILURE]		False
		lFiles = []
		lFiles_tmp = []
	
		if iOption.dir :
			lFiles_tmp = os.listdir( iOption.dir )
			for sFiles_tmp in lFiles_tmp :
				if os.path.isdir( sFiles_tmp ) :
					continue
				lFiles.append( sFiles_tmp )
			sChPath = iOption.dir
		elif iOption.file :
			lFiles.append( iOption.file )
			sChPath = os.path.dirname( iOption.file )
		else :
			return False
		
		if lFiles == [] :
			return False
		
		os.chdir( sChPath )
		
		return lFiles
	def fnSeperateFile(self, lFiles):
		# [IN] 		lFiles
		# [OUT]		List
		List = CLIST()
		for sfname in lFiles :
			sFormat = self.fnIsFormat( List.lKnown, sfname )
			if sFormat == False :
				sFormat = "Unknown"
			
			eval("List.l%s" % sFormat).append(sfname)
		return List
	def fnIsFormat(self, lKnown, sfname):
		# [IN] 		lKnown
		# [IN] 		sbuffer
		# [OUT]		sFormat 
		sFormat = False 
		for sKnown in lKnown :
			sFormat = eval("fnIs%s" % sKnown)(sfname)
			if sFormat != False :
				break
		return sFormat
		
class COPTIONS():
	def __del__(self):
		return
	def fnClassify(self, iOption, List):
		# [IN]		iOption
		# [IN]		List
		for sFormat in List.lKnown :
			if eval("List.l%s" % sFormat) == [] :
				continue
			
			if iOption.file :
				srcpath = os.path(iOption.file)
			elif iOption.dir :
				srcpath = iOption.dir
			else :
				return
			
			dstpath = srcpath + "\\%s" % sFormat
			if not os.path.exists(dstpath) :
				# create directory 
				os.mkdir(dstpath)
			else :
				if not os.path.isdir( dstpath ) :
					print "[Error] Path & File Name is Same! :: %s" % sFormat
					continue
			
			for sfname in eval("List.l%s" % sFormat) :
				if not CFile.fnMoveFile(srcpath, dstpath, sfname) :
					print "[Failure] fnMoveFile( ) :: %s" % sfname
		return 
	def fnScan(self, iOption, List):
		# [IN]		iOption
		# [IN]		List
		SCAN = CSCAN()
		for sFormat in List.lKnown :
			if eval("List.l%s" % sFormat) == [] :
				continue
			
			for sfname in eval("List.l%s" % sFormat) :
				print "[+] %s -" % sfname,
				if not eval("SCAN.fnScan%s" % sFormat)(sfname) :
					List.lExcept.append(sfname)
		
		if List.lExcept != [] :
			print "Except : %s" % str(List.lExcept)
		return
		
class CSCAN():
	def fnScanOLE(self, sfname):	# Office / HWP 5.x
		sbuffer = CFile.fnReadFile(sfname)
		
		# print "\t[Step 0] Header"
		tHeader = fnMap("OLEHeader", sbuffer[:SIZE_OLE_HEADER], RULE_OLE_HEADER_NAME, RULE_OLE_HEADER_PATTERN)
		
		# print "\t[Step 1] MSAT"
		tMSAT = CSECTOR.fnReadMSAT(sbuffer)
		if tMSAT == False :
			print "[Failure] fnReadMSAT()"
			return False
		
		# print "\t[Step 2] SAT"
		tSAT = CSECTOR.fnReadSAT(sbuffer, tMSAT)
		if tSAT == False :
			print "[Failure] fnReadSAT()"
			return False
		
		# print "\t[Step 3] Directory"
		sDirectory = CSECTOR.fnReadSector(sbuffer, tSAT, tHeader.DirSecID)
		
		nPos = 0
		lDirectory = []
		while nPos < len(sDirectory) :
			tDirectory = fnMap("OLEDirectory", sDirectory[nPos:nPos + SIZE_OLE_DIRECTORY], RULE_OLE_DIRECTORY_NAME, RULE_OLE_DIRECTORY_PATTERN)
			lDirectory.append( tDirectory )
			nPos += SIZE_OLE_DIRECTORY
		
		# print "\t[Step 4] SSAT"
		if tHeader.SSATSecID > 0 :
			tSSAT = CSECTOR.fnReadSSAT(sbuffer, tSAT, tHeader.SSATSecID)
			if tSSAT == False :
				print "[Failure] fnReadSSAT()"
				return False
		else :
			tSSAT = ()
			
		del sbuffer
		
		# print "\t[Step 5] Seperate OLE"
		Exploit = CExploit()
		
		# HWP 5.x
		for tDirectory in lDirectory :
			sDirName = CUtil.fnExtractAlphaNumber(tDirectory.DirName)
			if sDirName.find("Hwp") != -1 :
				if not Exploit.fnExploitHWP(sfname, lDirectory, tSAT, tSSAT) :
					return False
				return True
		
		# Office
		lOffice = ["DOC", "XLS", "PPT"]
		lOfficeSig = ["WordDocument", "Workbook", "PowerPointDocument"]
		
		for tDirectory in lDirectory :
			sDirName = CUtil.fnExtractAlphaNumber(tDirectory.DirName)
			if sDirName in lOfficeSig :
				sFormat = dict(zip(lOfficeSig, lOffice))[sDirName]
				break
		
		if not eval("Exploit.fnExploit%s" % sFormat)(sfname, lDirectory, tSAT, tSSAT) :
			return False
		return True	
	def fnScanPDF(self, sfname):
		try :
			from VScan.parse.Decode import CDecode
		except :
			print "[Error] Do not SCAN PDF File"
			return False
	
		# print "[Step1] Object"
		lObject = CREAD.fnObject(sfname)
		if lObject == False :
			print "[Failure] CREAD.fnObject( )"
			return False
		
		# print "[Step2] JavaScript"
		lJavaScript = []
		for tObject in lObject :
			# None-Decrypt : /Filter was't Exist but /JS or /JavaScript Exists
			if CUtilPDF.fnIsJavaScript(tObject[2]) :
				tObject[2] = CUtil.fnReplace(tObject[2])
				lJavaScript.append(tObject)
				continue
		
			# Need-Decrypt : /Filter must be Exist
			if tObject[1].find("/Filter") == -1 :
				continue
			
			lFilter = []
			for nIndex in range(len(DECODE_FILTER)-1) :
				if tObject[1].find( DECODE_FILTER[nIndex] ) != -1 or tObject[1].find( DECODE_FILTER_AKA[nIndex] ) != -1 :
					lFilter.append( DECODE_FILTER[nIndex] )
			if lFilter == [] :
				print "[Failure] Do not Found Decode"
				return False
			
			dParam = dict(zip(DECODE_FILTER, DECODE_FILTER_PARAM))
			lParam = []
			for sFilter in lFilter :
				if dParam[sFilter] == "No" :
					lParam.append( "None" )
					continue
				
				lPredict = CREAD.fnPredict(tObject[1])
				if lPredict == False :
					print "[Failure] CREAD.fnPredict( %s )" % tObject[0]
					return False
				
				lParam.append( eval("CDecode.fnParam%s" % sFilter)(tObject[1]) )
			
			sBody = tObject[2]
			for sfilter in enumerate(lFilter) :
				sBody = eval("CDecode.fn%s" % sfilter[1])(sBody, lParam[sfilter[0]] )
				if sBody == "" :
					print "[Error] Do not Support Decode Algorithm! :: %s" % sfilter
					return False
				elif sBody == False :
					print "[Failure] CDecode.fn%s( )" % sfilter
					return False
			
			if CUtil.fnIsJavaScript(sBody) :
				tObject[2] = fnReplace(sBody)
				lJavaScript.append(tObject)
			
		if lJavaScript == [] :
			print "None"
			return True
		
		# print "[Step3] Exploit"
		Exploit = CExploit()
		for tObject in lJavaScript :
			if not Exploit.fnExploitPDF(sfname, tObject) :
				print "[Failure] fnExploitPDF( ) :: %s" % tObject[0]
		return True
	def fnScanRTF(self, sfname):
		print "\n"
		return False
	def fnScanHWP(self, sfname):	# HWP 2.x / 3.x
		print "\n"
		return False	

class CExploit():
	def fnExploitDOC(self, sfname, lDirectory, SAT, SSAT):	
		try :
			from VScan.parse.DOC 	import CDOC
			from VScan.exploit.DOC 	import fnExploitDOC
		except :
			print "[ERROR] Do not SCAN ShellCode in DOC File"
			return True
	
		sbuffer = CFile.fnReadFile(sfname)
		sWordDocument = CSECTOR.fnReadByDirectory(sbuffer, lDirectory, SAT, SSAT, "WordDocument")
		if sWordDocument == False :
			print "[Failure] fnReadByDirectory( WordDocument )"
			return False
		
		nFib, nAdd = CDOC.fnFib(sWordDocument)
		
		# FibRgFcLcbBlob
		nPos = 0x5be + nAdd
		lFibRgFcLcb = []
		dFib = dict(zip(lVer, lFib))
		for nVer in lVer :
			lFibRgFcLcb.append(fnMap("FibRgFcLcb%s" % nVer, sWordDocument[nPos:nPos + eval("SIZE_FIBRGFCLCB_%s" % nVer)], eval("RULE_OFFICE_FIBRGFCLCB_NAME_%s" % nVer), eval("RULE_OFFICE_FIBRGFCLCB_PATTERN_%s" % nVer)))
			if nFib == dFib[nVer] :
				break
		
		# Table Stream
		sTblName = CDOC.fnTableName(sWordDocument)
		sTable = CSECTOR.fnReadByDirectory(sbuffer, lDirectory, SAT, SSAT, sTblName)
		if sTable == False :
			print "[Failure] fnReadByDirectory( %s )" % sTblName
			return False
		
		# Data Stream
		sData = CSECTOR.fnReadByDirectory(sbuffer, lDirectory, SAT, SSAT, "Data")
		if sData == False :
			print "[Failure] fnReadByDirectory( Data )"
			return False
		
		del sbuffer
		# Exploit
		bRet = fnExploitDOC(lFibRgFcLcb, sTable, sData)
		if bRet == True :
			print "None"
		
		return True
	def fnExploitXLS(self, sfname, lDirectory, SAT, SSAT):
		try :
			from VScan.parse.XLS 	import CXLS
			from VScan.exploit.XLS 	import fnExploitXLS
		except :
			print "[ERROR] Do not SCAN ShellCode in XLS File"
			return True
	
	
		print "Not Yet"
		return True
	def fnExploitPPT(self, sfname, lDirectory, SAT, SSAT):
		try :
			from VScan.parse.PPT 	import CPPT
			from VScan.exploit.PPT 	import fnExploitPPT
		except :
			print "[ERROR] Do not SCAN ShellCode in PPT File"
			return True
			
			
		print "Not Yet"
		return True
		
	def fnExploitHWP(self, sfname, lDirectory, SAT, SSAT):
		try :
			from VScan.exploit.HWP import fnExploitHWP
		except :
			print format_exc()
			print "[ERROR] Do not SCAN ShellCode in HWP 5.x File"
			return True
		
		sbuffer = CFile.fnReadFile(sfname)
		
		# print "\t[Step 5-1] RootEntry"
		sRootEntry = CSECTOR.fnReadByDirectory(sbuffer, lDirectory, SAT, SSAT, "RootEntry")
		if sRootEntry == False:
			print "[Failure] fnReadByDirectory( RootEntry )"
			return False
		
		# print "\t[Step 5-2] Section
		sSection = CSECTOR.fnReadByDirectory(sbuffer, lDirectory, SAT, SSAT, "Section")
		if sSection == False :
			print "[Failure] fnReadByDirectory( Section )"
			return False
		
		bInfect = False
		# Search HWP BOF
		# bRet = fnExploitHWP(sSection[:tDirectory.StreamSize])
		bRet = fnExploitHWP(sSection)
		if bRet == "Infect" :
			bInfect = True
		elif bRet == False :
			return False		
			
		# Search HWP Embedded PE
		for tDirectory in lDirectory :
			sDirName = CUtil.fnExtractAlphaNumber(tDirectory.DirName)
			if sDirName in ("RootEntry", "FileHeader", "Section0", "Section1") :
				continue
			
			sSector = CSECTOR.fnReadByDirectory(sRootEntry, lDirectory, SAT, SSAT, sDirName)
			if fnIsPE(sSector) == "PE" :
				print "[FOUND] Embedded PE : %s" % sDirName
				bInfect = True
		
		if bInfect == False :
			print "None"
				
		return True
		
	def fnExploitPDF(self, sfname, tObject):
		print "Not Yet"
		return False
	def fnExploitRTF():
		print "Not Yet"
		return False
	def fnExploitHWP23():
		print "Not Yet"
		return False

if __name__ == "__main__" :		
	main()
else :
	print "[-] Do not Support EXPORT!"
exit(0)

	

