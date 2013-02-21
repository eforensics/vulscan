#!/usr/bin/python
# -*- coding: utf-8 -*-

'''
author: Computists
website: http://computists.tistory.com
last modified: Feb 2013
'''

import sys, string
from struct import *
from collections import namedtuple
from WordStructure import *
from OleStructure import *

class OLE():
	def __init__(self, paths, log):
		self.filePath = paths
		self.log = log

	def MakeLogging (self, string):
		if self.log == None:
			print(string)
		else:
			self.log.WriteText(string)

	def OleParsing (self):
		filename = self.filePath
		hFile = open(filename,'rb')
		fileData = hFile.read()
		hFile.close()
		oleHeaderData = fileData[0:SIZE_HEADER]

		#=== get OLE Header ===#
		self.MakeLogging("[+] Get OLE Header Data\n")
		ntHeader = namedtuple('Header',ruleName_Header)
		sHeader = ntHeader._make(unpack(ruleChar_Header,oleHeaderData))

		#=== defined by OLE header ===#
		SECSIZE = 2 ** sHeader.SSZ
		SSECSIZE = 2 ** sHeader.SSSZ
		MINSECSIZE = sHeader.MiniSectorCutoff
		self.MakeLogging("\t[-] Done - OLE Header\n")

		#== get the Master Sector Allocation Table ==#
		self.MakeLogging("[+] Get Mater Sector Allcation Table\n")
		MSATTable = getMSAT(fileData,sHeader.MSAT, sHeader.SecCount_MSAT, sHeader.SecID_MSATStart)
		self.MakeLogging("\t[-] Done - MSAT table\n")

		#== get the Sector Allocation Table ==#
		self.MakeLogging("[+] Get Sector Allcation Table\n")
		SATTable = getSAT(fileData,sHeader.SecCount_SAT, MSATTable, SECSIZE)
		if SATTable == False:
			return False, False
		else:
			self.MakeLogging("\tDone - SAT table\n")

		#== get the Short-Sector Allocation Table ==#
		self.MakeLogging("[+] Get Short-Sector Allcation Table\n")
		SSATTable = getSSAT(fileData, sHeader.SecID_SSATStart, sHeader.SecCount_SSAT, SATTable, SECSIZE)
		if SSATTable == False:
			return False, False
		
		self.MakeLogging("\t[-] Done - SSAT table\n")

		#== get the Directory Entry information ==#
		self.MakeLogging("[+] Get Short-Sector Allcation Table\n")
		totalDir, DirCount = followSecs(SATTable,sHeader.SecID_DirStart,fileData, SECSIZE)
		if totalDir == False and DirCount == False:
			return False, False

		ntDirEntry = namedtuple('DirEntry',ruleName_DirEntry)
		aDirEntry = []

		for i in range((DirCount*4)) :
			aDirEntry.append(ntDirEntry._make(unpack(ruleChar_DirEntry,totalDir[i*SIZE_DIRENTRY:SIZE_DIRENTRY+i*SIZE_DIRENTRY])))

		self.MakeLogging("\t[-] Done - Directory Entry\n")

		#== get the Streams from each Directory ==#
		self.MakeLogging("[+] Get Stream Data from DirectoryEntry\n")

		SSecData, SSecCount = followSecs(SATTable,sHeader.SecID_SSATStart+1,fileData, SECSIZE)
		if SSecData == False and SSecCount == False:
			return False, False
		streamData = []

		for i in range(len(aDirEntry)) :
			if aDirEntry[i].StreamSize >= MINSECSIZE \
				or aDirEntry[i].DirName.replace('\x00','').find('Root Entry') != -1 \
				or aDirEntry[i].Type == 5:
				dirData, dirSecCount = followSecs(SATTable,aDirEntry[i].SecID_Start,fileData, SECSIZE)
				if dirData == False and dirSecCount == False:
					return False, False
			elif aDirEntry[i].StreamSize == 0: ### 디렉토리 사이즈가 0 일 때는 공백을 준다.
				dirData = ''				
			else:
				dirData, dirSecCount = followSSecs(SSATTable,aDirEntry[i].SecID_Start,SSecData, SSECSIZE)
				if dirData == False and dirSecCount == False:
					return False, False
			streamData.append(dirData[:aDirEntry[i].StreamSize])	
		self.MakeLogging("\t[-] Done - Stream Data\n")
	
		#=== OLE DirInfo ===#
		dirsInfo = {}
		for i in range(len(aDirEntry)) :
			dirInfo = (
						str(i),												### Number
						aDirEntry[i].DirName.replace('\x00',''),			### Directory name
						str(aDirEntry[i].Type),								### Directory Type
						str(aDirEntry[i].LeftChild).replace('-1','None'),	### Left Child
						str(aDirEntry[i].RightChild).replace('-1','None'),	### Right Child
						str(aDirEntry[i].RootChild).replace('-1','None'),	### Root Child
						str(aDirEntry[i].SecID_Start),						### SectionID
						str(aDirEntry[i].StreamSize)						### Directory size
						)
			
			dirsInfo[i] = dirInfo ### tuple 삽입
			print(aDirEntry[i].DirName)
			if aDirEntry[i].DirName.replace('\x00','').find('WordDocument') != -1:
				self.wordDocumentStream = streamData[i]
		
		return dirsInfo, streamData

	def WordParsing (self):
		sFIB = getFIB(self.wordDocumentStream)

		return sFIB

if __name__ == '__main__':
	filename = sys.argv[1]
	ole = OLE(filename, None)
	data, stream = ole.OleParsing()
	if data == False and stream == False:
		None
	else:
		wordData = ole.WordParsing()