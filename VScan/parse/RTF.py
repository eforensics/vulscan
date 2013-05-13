# External Module Import
import os

# Internal Module Import
from Common import CFile

def fnIsRTF(sData):
	# [IN]		sData
	# [OUT]		"RTF" / False
	RTFSig = ""
	try :
		if os.path.exists(sData) :
			RTFSig = CFile.fnReadFile(sData)[:6]
	except :
		RTFSig = sData[:6]
		
	if RTFSig == "{\rtf1" or RTFSig.find( "rt" ) != -1 :
		return "RTF"
	return False
		
	