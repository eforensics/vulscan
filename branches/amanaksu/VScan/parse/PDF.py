# External Module Import
import re, os
from traceback import format_exc

# Internal Module Import
from Common import CFile
from Decode import CDecode

def fnIsPDF(sData):
	# [IN]		sData
	# [OUT]		"PDF" / False
	sbuffer = ""
	try : 
		if os.path.exists(sData) :
			sbuffer = CFile.fnReadFile(sData)[:200]
	except :
		sbuffer = sData[:200]
			
	if bool( re.match("^%PDF-[0-9]\.[0-9]", sbuffer) ) :
		return "PDF"
	return False

class CUtilPDF():
	@classmethod
	def fnReplace(cls, sbuffer) :
		sbuffer = sbuffer.replace("\(", "(")
		sbuffer = sbuffer.replace("\)", ")")
		sbuffer = sbuffer.replace("\n", " ")
		sbuffer = sbuffer.replace("\r", " ")
		sbuffer = sbuffer.replace("\t", " ")
		sbuffer = sbuffer.replace("\\\\", "\\")	
		return sbuffer
	@classmethod
	def fnIsJavaScript(cls, sbuffer) :
		lJavaScriptStrings = ['var ',';',')','(','function ','=','{','}','if ','else','return','while ','for ',',','eval']
		lKeyStrings = [';','(',')']
		lFound = []
		nLimit = 15
		nResult = 0
		
		for chr in sbuffer :
			if ord(chr) >= 150 :
				return False
		
		for sJavaScript in lJavaScriptStrings :
			nCnt = sbuffer.count(sJavaScript)
			nResult += nCnt
			if nCnt > 0 :
				lFound.append(sJavaScript)
			elif nCnt == 0 and sJavaScript in lKeyStrings :
				return False
				
		if nResult > nLimit and len(lFound) >= 5 :
			return True
		return False
	
class CREAD():
	@classmethod
	def fnObject(cls, sfname) :
		sbuffer = CFile.fnReadFile(sfname)
		lObject = re.findall( RE_OBJECT, sbuffer, re.DOTALL )
		if lObject == [] :
			return False
		return lObject
	@classmethod
	def fnPredict(cls, sbuffer) :
		lPredict = re.findall(RE_PREDICT, sbuffer, re.DOTALL)
		if lPredict == [] :
			return False
		return lPredict
	
RE_OBJECT 	= "([0-9]+ [0-9] obj)(\s{0,}<<.{0,100},*?>>)[\s%]*?stream(.*?)endstream[\s%]*?endobj"
RE_PREDICT 	= ".{0,100}?/Predictor\s([0-9]{1,8})"
DECODE_FILTER = ["FlateDecode", "ASCII85Decode", "ASCIIHexDecode", "RunLengthDecode", "LZWDecode", "CCITTFaxDecode", "DCTDecode"]
DECODE_FILTER_AKA = ["Fl", "A85", "AHx", "RL", "LZW", "CCF", "DCT"]
DECODE_FILTER_PARAM = ["Yes", "No", "No", "No", "Yes", "Yes", "Yes"]

