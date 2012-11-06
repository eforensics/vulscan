# -*- coding:utf-8 -*-

# import Public Module
import traceback, os, binascii, zlib

# import Private Module
import Common
from PEScanner import PE

#------------------------------------------------------------------------------------------------
# [ Description ] 
# - HWP 2.0 or 3.0

class HWP23():
    @classmethod
    def Check(cls, pBuf):
        try :
            HWP23Sig = Common.BufferControl.Read(pBuf, 0, 0x20)
            if (HWP23Sig.find( "HWP Document File V2.00" ) != -1) or (HWP23Sig.find( "HWP Document File V3.00") != -1) :
                return "HWP23"
            else :
                return ""
        
        except :
            print traceback.format_exc()
    

    @classmethod
    def Scan(cls, File):        
        try :
            
            return True
        
        except :
            print traceback.format_exc()
            return False
            

#------------------------------------------------------------------------------------------------
# [ Description ]
# - HWP 5.x ( Compound File Format - OLE )

class HWP():
    @classmethod
    def Scan(cls, File):
        try :            
            Hwp = OperateHWP()           
            
            # Extract Sector Data
            if not Hwp.ExtractSector( File ) : 
                return False
            
            # Mapped File Header for Compression and Password
            MapHWP = MappedHWP()
            FileHeader = File["FileHeader"]
            if FileHeader == "" :
                print "\t\t[ERROR] FileHeader is NOT Saved"
                return False
            
            MapFHeader = MapHWP.MapFileHeader( FileHeader )
            if MapFHeader["Property"] & 1 :
                File["Compression"] = True
            else :
                File["Compression"] = False
                
            if MapFHeader["Property"] & 2 :
                File["Password"] = True
                print "\t\t[ERROR] Password is required"
                return False
            else :
                File["Password"] = False
            
            
            # Trace Sector Data
            Record = ParseRecord()
            if File["ExtType"] != [] :                
                DecType = Record.DecryptDataRecord(File["ExtType"])
                if DecType == [] :
                    print "\t\t[ERROR] DecryptDataRecord( )"
                    return False
                File["DecType"] = DecType
            
            
            # Checking Vulnerability
            # Case 1. Size Overflow ( 0xFFFFFFFF ) 
            if not Record.ParseDataRecord(File) :
                return False
            
            # Case 2. Embedded PE ( Just "MZ" )
            if not Record.CheckPEStream(File) :
                return False
            
            
        except :
            print traceback.format_exc()
            return False
        
        return True
    

class MappedHWP():
    def MapFileHeader(self, FileHeader):
        try :
            FHeader = {}
            OutFHeader = {}
            
            index = 0
            Position = 0
            for index in range( len(szFHeader) ) :
                if szFHeader[index] == 32 : 
                    Signature = Common.BufferControl.Read(FileHeader, Position, szFHeader[index])
                    FHeader[ mFHeader[index] ] = Common.BufferControl.ExtractAlphaNumber( Signature )
                elif szFHeader[index] == 4 :
                    FHeader[ mFHeader[index] ] = Common.BufferControl.ReadDword(FileHeader, Position)
                else :
                    FHeader[ mFHeader[index] ] = binascii.b2a_hex( Common.BufferControl.Read(FileHeader, Position, szFHeader[index]) )
                
                Position += szFHeader[index]
            
        except :
            print traceback.format_exc()
            return OutFHeader
        
        OutFHeader = FHeader
        return OutFHeader
    
#----------------------------------------------------------------------------------------------------------------------------------------------------#
#    FileHeader ( HWP )
#----------------------------------------------------------------------------------------------------------------------------------------------------#
szFHeader = [32,4,4,216]
                            #   Offset   : Size : Description
mFHeader = ["Signature",    # 000 [0x00] : 0x20 : Signature "HWP Document File"
            "Version",      # 032 [0x20] : 0x04 : File Version. 0xMMnnPPrr ( Ex : 5.0.3.0 )
            "Property",     # 036 [0x24] : 0x04 : Property
            "Reserved"]     # 040 [0x28] : 0xD8 : Reserved

"""
[ Property ]
Bit       Description
12 ~ 31   Reserved
11        CCL Document ( Enable / Disable )
10        Accredited certificate DRM Security Document ( Enable / Disable )
9         Electronic signature spare storage ( Enable / Disable )
8         Accredited certificate encryption ( Enable / Disable )
7         Electronic Signature ( Enable / Disable )
6         Traceability ( Enable / Disable )
5         XML Template Storage  ( Enable / Disable )
4         DRM Security Document ( Enable / Disable )
3         Saved Script ( Enable / Disable )
2         for Distribution ( Enable / Disable )
1         Password ( Set / UnSet )     =-> Needs
0         Pack ( Enable / Disable )    =-> Needs
"""
#----------------------------------------------------------------------------------------------------------------------------------------------------#

    
class OperateHWP():   
    def ExtractSector(self, File):
        try : 
            pBuf = File["pBuf"]
            DirList = File["DirList"]
            ExtType = []
            
            # Extract Sector By SAT
            SATable = File["SATable"]
            RefSAT = File["RefSAT"]
            for RefEntry in RefSAT :
                for DirEntry in DirList :
                    if RefEntry == DirEntry["EntryName"] :
                        Sector = self.ExtractSectorbySAT(pBuf, SATable, DirEntry["SecID"], DirEntry["szData"])
                        fname = "%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"])
                        Common.FileControl.WriteFile(fname, Sector)
                        
                        if RefEntry == "RootEntry" :
                            File["RootEntry"] = Sector                        
                        
                        if RefEntry == "FileHeader" :
                            File["FileHeader"] = Sector
                        
                        if RefEntry == "DocInfo" or (RefEntry.find("BinaryData") != -1) or (RefEntry.find("VersionLog") != -1) or (RefEntry.find("Section") != -1) :
                            ExtType.append(fname)
                        
                        break
                
            # Extract Sector By SSAT
            pBuf = File["RootEntry"]
            if pBuf == "" :
                print "\t\t[ERROR] RootEntry is NOT Saved"
                return False
            
            SSATable = File["SSATable"]
            RefSSAT = File["RefSSAT"]
            for RefEntry in RefSSAT :
                for DirEntry in DirList :
                    if RefEntry == "" :
                        continue                    
                        
                    if RefEntry == DirEntry["EntryName"] :
                        Sector = self.ExtractSectorbySSAT(pBuf, SSATable, DirEntry["SecID"], DirEntry["szData"])
                        fname = "%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"])
                        Common.FileControl.WriteFile(fname, Sector)
                        
                        if RefEntry == "FileHeader" :
                            File["FileHeader"] = Sector
                            
                        if RefEntry == "DocInfo" or (RefEntry.find("BinaryData") != -1) or (RefEntry.find("VersionLog") != -1) or (RefEntry.find("Section") != -1) :
                            ExtType.append(fname)
                            
                        break
            
            File["ExtType"] = ExtType
            
        except :
            print traceback.format_exc()
            return False
            
        return True


    
    def ExtractSectorbySAT(self, pBuf, Table, SecID, Size):
        try :
            Sector = ""
            tmpSector = ""
            while True :
                if SecID == 0xffffffff or SecID == 0xfffffffe or SecID == 0xfffffffd or SecID == 0xfffffffc :
                    break
            
                tmpSector += Common.BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x200)
                SecID = Table[SecID]
            
            Sector = Common.BufferControl.Read(tmpSector, 0, Size)
            
        except :
            print traceback.format_exc()
            return ""
            
        return Sector


    def ExtractSectorbySSAT(self, pBuf, Table, SecID, Size):
        try :
            Sector = ""
            tmpSector = ""
            while True :
                if SecID == 0xffffffff or SecID == 0xfffffffe or SecID == 0xfffffffd or SecID == 0xfffffffc :
                    break
            
                tmpSector += Common.BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x40)
                SecID = Table[SecID]
            
            Sector = Common.BufferControl.Read(tmpSector, 0, Size)
            
        except :
            print traceback.format_exc()
            return ""
        
        return Sector


class ParseRecord():
    def DecryptDataRecord(self, ExtType):
        try :
            DecType = []
            for fname in ExtType :            
                Buf = Common.FileControl.ReadFileByBinary(fname)
                Buf_Dec = zlib.decompress(Buf, -15)
                sname = "%s_Decrypt" % fname
                Common.FileControl.WriteFile(sname, Buf_Dec)
                if (sname.find("DocInfo") != -1) or (sname.find("BinaryData") != -1) or (sname.find("VersionLog") != -1) or (sname.find("Section") != -1):
                    DecType.append(sname)
                else :
                    print "Does not belog anywhere! ( %s )" % sname
            
        except :
            print traceback.format_exc()
            return []
        
        return DecType
    
    
    def ParseDataRecord(self, File):
        try :
            DecType = File["DecType"] 
            
            # Parse Decrypted Type
            for fname in DecType :
                
                print "\t    %s" % fname
                
                Position = 0
                DecrpytBuf = Common.FileControl.ReadFileByBinary(fname)
                while Position < len( DecrpytBuf ) :
                    OverFlag = False
                    DataRecord = Common.BufferControl.ReadDword(DecrpytBuf, Position)
                    
                    Position += 4
                    TagID = DataRecord & 0x000003FF
                    Level = ( DataRecord & 0x000FFC00 ) >> 10
                    Size = ( DataRecord & 0xFFF00000 ) >> 20
                    if Size == 0xFFF :
                        OverFlag = True
                        Size = Common.BufferControl.ReadDword(DecrpytBuf, Position)
                        Position += 4
                    
                    if (OverFlag == True and Size > 0x00010000) or (OverFlag == True and Size == 0xFFFFFFFF) :
                        print "\t\t\t0x%08X : ( 0x%08X   %03X   0x%08X (%s) ) - %s   " % ( Position, TagID, Level, Size, OverFlag, HWPTAG[TagID]) \
                              + ">" * 10 + "Suspicious"
                        
                    Position += Size
                    
        except :
            print traceback.format_exc()
            return False
        
        return True


    def CheckPEStream(self, File):
        try :
            dirpath = os.path.abspath( os.curdir )
            
            # Extract File ( "fname * .dump" )
            TarList = []
            flist = os.listdir( dirpath )
            name = os.path.splitext( File["fname"] )
            for fname in flist :
                if fname.find(name[0]) != -1 and fname.find(".dump") :
                    TarList.append( fname )
            
            
            # Check PE 
            CheckFormat = PE()
            for fname in TarList :
                pBuf = Common.FileControl.ReadFileByBinary( fname )
                if CheckFormat.Check(pBuf) :
                    print "\t    [-] Find PE : %s" % fname
            
        except :
            print traceback.format_exc()
            return False
        
        return True



NormalStream = ["BodyText", "DefaultScript", "DocInfo", "DocOptions", "FileHeader", "HwpSummaryInformation", "JScriptVersion", "LinkDoc", "PrvImage", "PrvText", "RootEntry"]

#----------------------------------------------------------------------------------------------------------------------------------------------------#
#    TagID 
#----------------------------------------------------------------------------------------------------------------------------------------------------#
HWPTAG_BEGIN = 0x010
        # DocInfo
HWPTAG = {HWPTAG_BEGIN      :"HWPTAG_DOCUMENT_PROPERTIES",
          HWPTAG_BEGIN+1    :"HWPTAG_ID_MAPPINGS",
          HWPTAG_BEGIN+2    :"HWPTAG_BIN_DATA",
          HWPTAG_BEGIN+3    :"HWPTAG_FACE_NAME",
          HWPTAG_BEGIN+4    :"HWPTAG_BORDER_FILL",
          HWPTAG_BEGIN+5    :"HWPTAG_CHAR_SHAPE",
          HWPTAG_BEGIN+6    :"HWPTAG_TAB_DEF",
          HWPTAG_BEGIN+7    :"HWPTAG_NUMBERING",
          HWPTAG_BEGIN+8    :"HWPTAG_BULLET",
          HWPTAG_BEGIN+9    :"HWPTAG_PARA_SHAPE",
          HWPTAG_BEGIN+10   :"HWPTAG_STYLE",
          HWPTAG_BEGIN+11   :"HWPTAG_DOC_DATA",
          HWPTAG_BEGIN+12   :"HWPTAG_DISTRIBUTE_DOC_DATA",
          HWPTAG_BEGIN+13   :"RESERVED",
          HWPTAG_BEGIN+14   :"HWPTAG_COMPATIBLE_DOCUMENT",
          HWPTAG_BEGIN+15   :"HWPTAG_LAYOUT_COMPATIBILITY",
          
          # BodyText 
          HWPTAG_BEGIN+50   :"HWPTAG_PARA_HEADER",
          HWPTAG_BEGIN+51   :"HWPTAG_PARA_TEXT",
          HWPTAG_BEGIN+52   :"HWPTAG_PARA_CHAR",
          HWPTAG_BEGIN+53   :"HWPTAG_PARA_LINE_SEG",
          HWPTAG_BEGIN+54   :"HWPTAG_PARA_RANGE_TAG",
          HWPTAG_BEGIN+55   :"HWPTAG_CTRL_HEADER",
          HWPTAG_BEGIN+56   :"HWPTAG_LIST_HEADER",
          HWPTAG_BEGIN+57   :"HWPTAG_PAGE_DEF",
          HWPTAG_BEGIN+58   :"HWPTAG_FOOTNOTE_SHAPE",
          HWPTAG_BEGIN+59   :"HWPTAG_PAGE_BORDER_FILL",
          HWPTAG_BEGIN+60   :"HWPTAG_SHAPE_COMPONENT",
          HWPTAG_BEGIN+61   :"HWPTAG_TABLE",
          HWPTAG_BEGIN+62   :"HWPTAG_SHAPE_COMPONENT_LINE",
          HWPTAG_BEGIN+63   :"HWPTAG_SHAPE_COMPONENT_RECTANGLE",
          HWPTAG_BEGIN+64   :"HWPTAG_SHAPE_COMPONENT_ELLIPSE",
          HWPTAG_BEGIN+65   :"HWPTAG_SHAPE_COMPONENT_ARC",
          HWPTAG_BEGIN+66   :"HWPTAG_SHAPE_COMPONENT_POLYGON",
          HWPTAG_BEGIN+67   :"HWPTAG_SHAPE_COMPONENT_CURVE",
          HWPTAG_BEGIN+68   :"HWPTAG_SHAPE_COMPONENT_OLE",
          HWPTAG_BEGIN+69   :"HWPTAG_SHAPE_COMPONENT_PICTURE",
          HWPTAG_BEGIN+70   :"HWPTAG_SHAPE_COMPONENT_CONTAINER",
          HWPTAG_BEGIN+71   :"HWPTAG_CTRL_DATA",
          HWPTAG_BEGIN+72   :"HWPTAG_EQEDIT",
          HWPTAG_BEGIN+73   :"RESERVED",
          HWPTAG_BEGIN+74   :"HWPTAG_SHAPE_COMPONENT_TEXTART",
          HWPTAG_BEGIN+75   :"HWPTAG_FORM_OBJECT",
          HWPTAG_BEGIN+76   :"HWPTAG_SHAPE",
          HWPTAG_BEGIN+77   :"HWPTAG_MEMO_LIST",
          
          # DocInfo          
          HWPTAG_BEGIN+78   :"HWPTAG_FORBIDDEN_CHAR",
          
          # BodyText
          HWPTAG_BEGIN+79   :"HWPTAG_CHART_DATA",
          HWPTAG_BEGIN+99   :"HWPTAG_SHAPE_COMPONENT_UNKNOWN" }




