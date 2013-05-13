# External Module Import
import os
from traceback import format_exc

# Internal Module Import
from Common import CFile

def fnIsPE(sData):
	# [IN]		sfname
	# [OUT]		"PE" / False
	MZSig = ""
	try :
		if os.path.exists(sData) :
			MZSig = CFile.fnReadFile(sData)[:2]	
	except :
		MZSig = sData[:2]
	
	if MZSig == "MZ" or MZSig == "mz" :
		return "PE"
	return False

	