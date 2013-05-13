# External Module Import
import os
from traceback import format_exc

# Internal Module Import
from Common import CFile

def fnIsHWP(sData) :	# HWP 2.x / 3.x
	# [IN]		sData
	# [OUT]		"HWP" / False
	sbuffer = ""
	try :
		if os.path.exists(sData) :
			sbuffer = CFile.fnReadFile(sData)[:0x20]
	except :
		sbuffer = sData[:0x20]
		
	if (sbuffer.find("HWP Document File V2.00") != -1) or (sbuffer.find("HWP Document File V3.00") != -1) :
		return "HWP"
	return False
	