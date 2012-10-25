# -*- coding:utf-8 -*-

# import Public Module
import traceback, binascii, os, sys
import struct

from collections import namedtuple

# import private module
from Common import BufferControl, FileControl
from HWPScanner import HWP
from OfficeScanner import Office


class OLEStruct():   
    def __init__(self, File):
        # OLE Header
        File["NumSAT"] = ""
        File["DirSecID"] = ""
        File["SSATSecID"] = ""
        File["NumSSAT"] = ""
        File["NumMSAT"] = "" 
        File["MSAT"] = []
        
        # Table Information
        File["SATable"] = ""
        File["SATSectors"] = ""
        File["Directory"] = ""
        File["DirList"] = []
        
        File["RefSAT"] = []
        File["RefSSAT"] = []
        
        return
    
    
    def OLEHeader(self, File):
        try :
            # Parse Header 
            MapOLE = MappedOLE()
            Header = MapOLE.MapHeader(File)
                        
#            PrintOLE.PrintHeader( File, Header )
            
            # Parse MSAT
            MSAT = []
            pBuf = File["pBuf"]
            MSATSectors = Header["MSAT"]
            if Header["NumSAT"] > 109 or Header["NumMSAT"] > 0 :
                Position = Header["MSATSecID"] * 0x200 + 0x200
                Size = 0x200
                MSATSectors += BufferControl.Read(pBuf, Position, Size)
            
            MSATList = BufferControl.ConvertBinary2List(MSATSectors, 4)            
            
            SecID = 0
            for index in range( len(MSATList) ) :
                SecID = MSATList[index]
                MSAT.append( SecID )
                if SecID == 0xffffffff or SecID == 0xfffffffe or SecID == 0xfffffffd or SecID == 0xfffffffc :
                    break
                
                if SecID > len(MSATList) :
                    break
        
        except :
            print traceback.format_exc()

        
        # Exception Checking
        if (os.path.getsize( File["fname"] ) < 0x700000) and (Header["NumSAT"] > 109 or Header["NumMSAT"] > 0) :
            File['logbuf'] += "\n     Suspicious ( 0x%08X )" % os.path.getsize( File["fname"] )
        
#        if Header["NumSAT"] > 109 :   
#            File['logbuf'] += "\n    [Failure] Exist Extra SAT!!! ( 0x%08X )" % Header["NumSAT"]     
#            return False
#          
        # Saving Need Data
        File["NumSAT"] = Header["NumSAT"]
        
        File["DirSecID"] = Header["DirSecID"]
        
        File["SSATSecID"] = Header["sSATSecID"]
        File["NumSSAT"] = Header["NumsSAT"]
        
        File["NumMSAT"] = Header["NumMSAT"] 
        File["MSAT"] = MSAT
        
        return True            


    def OLETableSAT(self, File):
        try :
            # Get SAT Sectors
            Sector = ""
            index = 0
            MSAT = File["MSAT"]
            for index in MSAT :    
                if index == 0xffffffff or index == 0xfffffffe or index == 0xfffffffd or index == 0xfffffffc :
                    break
                Sector += BufferControl.ReadSectorByBuffer(File["pBuf"], index, 0x200)
                
            if Sector == "" :
                File['logbuf'] += "\n    [Failure] ReadSectorByBuffer( MSAT[0] : 0x%08X ) for OLETableSAT( )" % MSAT[0]
                return False             
            
            File["SATSectors"] = Sector 
            
            
            # Convert SAT Table Type
            SAT = BufferControl.ConvertBinary2List(Sector, 4)
            if SAT == [] :
                File['logbuf'] += "\n    [Failure] ConvertBinary2List( ) for OLETableSAT( )"
                return False

            File["SATable"] = SAT
            
        except :
            traceback.format_exc()
            return False
            
        return True
    

    def OLETableSSAT(self, File):
        try :
            # Get SSAT Sector
            Sector = ""
            SAT = File["SATable"]
            Sector = OLEStruct.OLETableTraceBySecID(File["fname"], File["pBuf"], SAT, File["SSATSecID"], 0x200)
            if Sector == "" :
                File['logbuf'] += "\n    [Failure] OLETableTraceBySecID( ) for OLETableSSAT( )"
                return False
            
            File["SSATSectors"] = Sector
            
            # Convert SSAT Table Type
            SSAT = BufferControl.ConvertBinary2List(Sector, 4)
            if SSAT == [] :
                File['logbuf'] += "\n    [Failure] ConvertBinary2List( ) for OLETableSSAT( )"
                return False
            
            File["SSATable"] = SSAT
            
        except :
            print traceback.format_exc()
            return False
            
        return True


    def OLEDirectory(self, File):
        try :
            # Dump Directory 
            SAT = File["SATable"]  
            Sectors = OLEStruct.OLETableTraceBySecID(File["fname"], File["pBuf"], SAT, File["DirSecID"], 0x200)
            if Sectors == "" :
                File['logbuf'] += "\n    [Failure] OLETableTraceBySecID( ) for OLEDirectory( )"
                return False

            File["Directory"] = Sectors


            # Procedure Directory
            MapOLE = MappedOLE()
            
            Cnt = 0
            DirCnt = len( Sectors ) / 0x80               
            DirList = []
            RefSAT = []
            RefSSAT = []
            
            while Cnt < DirCnt : 
                # Get Directory Sector Data
                DirData = BufferControl.Read(Sectors, Cnt * 0x80, 0x80)
                
                # Parse Directory
                Directory = MapOLE.MapDirectory(File, DirData)
                DirList.append( Directory )
                Cnt += 1
                
                # Separate Directory
                if Directory["EntryType"] == '05' or Directory["szData"] > 0x1000 :
                    RefSAT.append( Directory["EntryName"] )
                elif Directory["EntryType"] != '05' and Directory["szData"] <= 0x1000 :
                    RefSSAT.append( Directory["EntryName"] )
                else :
                    File['logbuf'] += "\n    [Failure] Failed Separation by referred type ( %s )" % Directory["EntryName"]
            
                
            # File Format Re-Separation
            File["format"] = ""
            for Directory in DirList :
                if Directory["EntryName"][:3].find( "Hwp" ) != -1 :
                    File["format"] = "HWP"
            
            if File["format"] == "" :
                File["format"] = "Office"
        
        
        except :
            traceback.format_exc()
            return False
        
        File["DirList"] = DirList
        File["RefSAT"] = RefSAT
        File["RefSSAT"] = RefSSAT
        return True

    
    @classmethod
    def OLETableTraceBySecID(cls, fname, pBuf, table, SecID, Size):
        try :
            Sector = ""
            while True :
                Sector += BufferControl.ReadSectorByBuffer(pBuf, SecID, Size)
                SecID = table[SecID]
                
                if SecID in ExceptSecID :
                    break
                
                if SecID > len(table) :
                    print "   %s ( Over SecID : %x / %x) - Suspicious" % (fname, SecID, len(table))
                    break

        except IndexError :
            print "   %s ( SecID : %x / szTable : %x) - IndexError" % (fname, SecID, len(table))
            
            wName = "%s_%x_TableIndexError.dump" % (fname, SecID)
            OutBuf = BufferControl.ConvertList2Binary( table )
            FileControl.WriteFile(wName, OutBuf)
                
        except :
            print traceback.format_exc()
        
        return Sector    

ExceptSecID = [0xffffffff, 0xfffffffe, 0xfffffffd, 0xfffffffc, 0xfeffffff, 0xfdffffff, 0xfcffffff]


class MappedOLE():
    def MapHeader(self, File):
        try :
            Header = {}
            OutHeader = {}
            pBuf = File["pBuf"]
            
            index = 0
            Position = 0
            for index in range( len(szHeader) ) :
                if szHeader[index] == 2 :
                    Header[ mHeader[index] ] = BufferControl.ReadWord(pBuf, Position)
                elif szHeader[index] == 4 :
                    Header[ mHeader[index] ] = BufferControl.ReadDword(pBuf, Position)
                elif szHeader[index] == 8 :     # for Signature 
                    Header[ mHeader[index] ] = binascii.b2a_hex( BufferControl.Read(pBuf, Position, szHeader[index]) )
                else :
                    Header[ mHeader[index] ] = BufferControl.Read(pBuf, Position, szHeader[index])
                
                Position += szHeader[index]
                    
        except :
            print traceback.format_exc()
            return OutHeader
    
        OutHeader = Header
        return OutHeader


    def MapDirectory(self, File, DirData):
        try :
            Directory = {}
            OutDirectory = {}
            
            index = 0
            Position = 0
            for index in range( len(szDirEntry) ) :
                if szDirEntry[index] == 2 :
                    Directory[ mDirEntry[index] ] = BufferControl.ReadWord(DirData, Position)
                elif szDirEntry[index] == 4 :
                    Directory[ mDirEntry[index] ] = BufferControl.ReadDword(DirData, Position)
                elif szDirEntry[index] == 64 :
                    EntryName = BufferControl.Read(DirData, Position, szDirEntry[index])
                    Directory[ mDirEntry[index] ] = BufferControl.ExtractAlphaNumber( EntryName )                        
                else :
                    Directory[ mDirEntry[index] ] = binascii.b2a_hex( BufferControl.Read(DirData, Position, szDirEntry[index]))
            
                Position += szDirEntry[index]
            
        except :
            print traceback.format_exc()
            return OutDirectory
        
        OutDirectory = Directory
        return OutDirectory
            

#----------------------------------------------------------------------------------------------------------------------------------------------------#
#    Header
#----------------------------------------------------------------------------------------------------------------------------------------------------#
szHeader = [8,16,2,2,2,2,2,10,4,4,4,4,4,4,4,4,436]
                                #  Offset   : Size  : Description
mHeader = ["Signature",     # 00 [0x00] : 0x008 : Compound Document File Identifier ( 0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1 )
           "UID",           # 08 [0x08] : 0x010 : Unique Identifier
           "Revision",      # 24 [0x18] : 0x002 : Revision Number ( Most 0x003E )
           "Version",       # 26 [0x1A] : 0x002 : Version Number ( Most 0x0003 )
           "OrderID",       # 28 [0x1C] : 0x002 : Order Identifier ( 0xFE 0xFF = Little-Endian / 0xFF 0xFE = Big-Endian )
           "szSector",      # 30 [0x1E] : 0x002 : Size of Sector
           "szShort",       # 32 [0x20] : 0x002 : Size of Short-Sector
           "NotUsed1",      # 34 [0x22] : 0x00A : Not Used
           "NumSAT",        # 44 [0x2C] : 0x004 : Total Number of Sectors used for the SAT
           "DirSecID",      # 48 [0x30] : 0x004 : SecID of First Sector of the Storage ( Directory ) Stream referred SAT
           "NotUsed2",      # 52 [0x34] : 0x004 : Not Used
           "szMinSector",   # 56 [0x38] : 0x004 : Minimum size of a standard stream
           "sSATSecID",     # 60 [0x3C] : 0x004 : SecID of First Sector of the Short-SAT
           "NumsSAT",       # 64 [0x40] : 0x004 : Total Number of Sectors used for the Short-SAT
           "MSATSecID",     # 68 [0x44] : 0x004 : SecID of First Sector of the MSAT
           "NumMSAT",       # 72 [0x48] : 0x004 : Total Number of Sectors used for the MSAT
           "MSAT"]          # 76 [0x4C] : 0x1B4 : First Part of the MSAT containing 109 SecIDs

#----------------------------------------------------------------------------------------------------------------------------------------------------#
#    Directory
#----------------------------------------------------------------------------------------------------------------------------------------------------#
szDirEntry = [64,2,1,1,4,4,4,16,4,8,8,4,4,4]
                            #   Offset   : Size : Description
mDirEntry = ["EntryName",   # 000 [0x00] : 0x40 : Character Array of the name of the entry ( Unicode )
             "szName",      # 064 [0x40] : 0x02 : Size of the used area of the Character buffer of the name
             "EntryType",   # 066 [0x42] : 0x01 : Type of Entry ( 00 = Empty, 01 = UserStorage, 02 = UserStream, 03 = LockBytes, 04 = Property, 05 = RootStorage
             "EntryNode",   # 067 [0x43] : 0x01 : Node Color of the Entry ( 00 = Red, 01 = Black 
             "LeftChild",   # 068 [0x44] : 0x04 : DirID of the Left Child
             "RightChild",  # 072 [0x48] : 0x04 : DirID of the Right Child
             "SubNode",     # 076 [0x4C] : 0x04 : DirID of the Sub Node Entry
             "UID",         # 080 [0x50] : 0x10 : Unique Identifier
             "UserFlag",    # 096 [0x60] : 0x04 : User flags
             "Time_Create", # 100 [0x64] : 0x08 : Time Stamp of Creation of this entry
             "Time_Modify", # 108 [0x6C] : 0x08 : Time Stamp of Modification of this entry
             "SecID",       # 116 [0x74] : 0x04 : SecID of First Sector or Short-Sector
             "szData",      # 120 [0x78] : 0x04 : Total Data Size in bytes
             "NotUsed"]     # 124 [0x7C] : 0x04 : Not Used

#----------------------------------------------------------------------------------------------------------------------------------------------------#

    
class PrintOLE():
    @classmethod
    def PrintHeader(cls, File, Header):
        File['logbuf'] += "\n\n\t" + "=" * 79
        File['logbuf'] += "\n\tHeader"
        File['logbuf'] += "\n\t" + "-" * 79
        File['logbuf'] += "\n\tSignature\t\t\t\t\t\t: %s" % Header['Signature']
        File['logbuf'] += "\n\tSector Size\t\t\t\t\t\t: 0x%08X" % (1<<Header['szSector'])
        File['logbuf'] += "\n\tShort-Sector Size\t\t\t\t: 0x%08X" % (1<<Header['szShort'])
        File['logbuf'] += "\n\tSAT Sector Count\t\t\t\t: 0x%08X" % Header['NumSAT']
        File['logbuf'] += "\n\tDirectory Sector's First SecID\t: 0x%08X" % Header['DirSecID']
        File['logbuf'] += "\n\tShort-SAT Sector's First SecID\t: 0x%08X" % Header['sSATSecID']
        File['logbuf'] += "\n\tShort-SAT Sector Count\t\t\t: 0x%08X" % Header['NumsSAT']
        File['logbuf'] += "\n\tMaster-SAT Sector's First SecID\t: 0x%08X" % Header['MSATSecID']
        if Header['NumMSAT'] > 0 :
            File['logbuf'] += "\n\tMaster-SAT Sector Count\t\t\t: 0x%08X -> Need to Big Data Proceduring" % Header['NumMSAT']
        else :
            File['logbuf'] += "\n\tMaster-SAT Sector Count\t\t\t: 0x%08X" % Header['NumMSAT']
        File['logbuf'] += "\n\t" + "=" * 79 + "\n"
    
    
    @classmethod
    def PrintDirectory(cls, File, Directory):
        EntryName = BufferControl.ExtractAlphaNumber( Directory["EntryName"] )
        File['logbuf'] += "\n\t    %s" % EntryName
        if len(EntryName) < 21 :
            space = 21 - len(EntryName)
            File['logbuf'] += " " * space
        
        if Directory["EntryType"] == '01' :
            File['logbuf'] += "\tStorage"                      
        elif Directory["EntryType"] == '02' : 
            File['logbuf'] += "\tStream"
        elif Directory["EntryType"] == '05' :
            File['logbuf'] += "\tRoot" 
        else :
            File['logbuf'] += "\tEmpty"
                
        File['logbuf'] += "\t0x%08x" % Directory["SecID"] 
        File['logbuf'] += "\t0x%08x" % Directory["szData"]
                
        if Directory["EntryType"] == '05' :
            File['logbuf'] += "\tReferred SAT"
        elif Directory["EntryType"] != '05' and Directory["szData"] > 0x1000 :
            File['logbuf'] += "\tReferred SAT"
        elif Directory["EntryType"] != '05' and Directory["szData"] <= 0x1000 :
            File['logbuf'] += "\tReferred Short-SAT"
        else :
            File['logbuf'] += "Error"  
    
    
    
    
class OLEScan():
    @classmethod
    def Check(cls, pBuf):
        try :
            # Case1. 0xe011cfd0, 0xe11ab1a1L
            if BufferControl.ReadDword(pBuf, 0) == 0xe011cfd0L and BufferControl.ReadDword(pBuf, 0x4) == 0xe11ab1a1L :
                return "OLE"
            # Case2. 0xe011cfd0, 0x20203fa1
            elif BufferControl.ReadDword(pBuf, 0) == 0xe011cfd0L :
                return "OLE"
            else :
                return ""
        except struct.error : 
            return ""
        
        except :
            print traceback.format_exc()
            
            
    @classmethod
    def Scan(cls, File):
        try :
            File['logbuf'] += "\n    [+] %s...........%s" % ( File["fname"], File["format"] )
            
            OLE = OLEStruct( File )
            
            if not OLE.OLEHeader(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLEHeader( %s )" % File["fname"]
                return False
    
            
            if not OLE.OLETableSAT(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLETableSAT( %s )" % File["fname"]
                return False
    
            if not OLE.OLETableSSAT(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLETableSSAT( %s )" % File["fname"]
    
            
            if not OLE.OLEDirectory(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLEDirectory( %s )" % File["fname"]
                return False

            DocScan = {"HWP":HWP.HWPScan, "Office":Office.OfficeScan}
            for field in DocScan :
                if File["format"] == field :
                    DocScan[ field ]( File ) 
            
        except :
            print traceback.format_exc()
            return False
        
        return True
    





