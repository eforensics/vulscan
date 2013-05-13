# External Module Import
from struct import unpack
from traceback import format_exc
from string import letters, digits
import shutil

class CFile():
	@classmethod
	def fnReadFile(cls, sfname) :
		# [IN]		sfname
		# [OUT]
		#	- [SUCCESS]		sbuffer
		#	- [ERROR]		None
		try :
			with open(sfname, "rb") as pFile :
				sbuffer = pFile.read()
		except :
			print format_exc()
			return None
		return sbuffer
	@classmethod
	def fnWriteFile(cls, sfname, sbuffer) :
		# [IN]		sfname
		# [IN]		sbuffer
		# [OUT]		True (make file) / None (error)
		try :
			with open(sfname, "wb") as pFile :
				try :
					pFile.write(sbuffer)
				except :
					pFile.write(str(sbuffer))
		except :
			print format_exc()
			return None
		return True
	@classmethod
	def fnCopyFile(cls, srcpath, dstpath, sfname) :
		srcfile = srcpath + "\\%s" % sfname
		dstfile = dstpath + "\\%s" % sfname
		try :
			shutil.copyfile(srcfile, dstfile)
		except :
			return False
		return True
	
class CBuffer():
	@classmethod
	def fnReadData(cls, sbuffer, nOffset, nSize) :
		# [IN]		sbuffer
		# [IN]		nOffset
		# [IN]		nSize
		# [OUT]		data / None (error)
		try :
			Sig = {1:"<B", 2:"<H", 4:"<L"}
			Val = sbuffer[nOffset:nOffset+nSize]
			return unpack(Sig[nSize], Val)[0]
		except :
			print format_exc()
			return None

class CUtil():
	@classmethod
	def fnExtractAlphaNumber(cls, sbuffer) :
		sOutBuffer = ""
		sPattern = letters + digits
		for sOneByte in sbuffer :
			if sOneByte in sPattern :
				sOutBuffer += sOneByte
		return sOutBuffer
	@classmethod
	def fnBitParse(cls, nFlag, nOffset, nCnt) :
		return (nFlag / 2**nOffset) & (2**nCnt -1)
	@classmethod
	def fnPack(cls, sbuffer, nSize) :
		dSize = {1:"b", 2:"h", 4:"l"}
		return unpack(str(len(sbuffer)/nSize) + dSize[nSize] , sbuffer)
