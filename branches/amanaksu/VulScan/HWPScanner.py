# -*- coding:utf-8 -*-

# import Public Module
import traceback, binascii, os, zlib


# import private module
from ComFunc import BufferControl, FileControl


class HWP():
    @classmethod
    def HWPScan(cls, File):
        try :
            File['logbuf'] += " : HWP"
            Hwp = OperateHWP()           
            
            # Extract Sector Data
#            File['logbuf'] += "\n        [Extract] "
            if not Hwp.ExtractSector( File ) : 
                return False
            
            
            # Mapped File Header for Compression and Password
            MapHWP = MappedHWP()
            FileHeader = File["FileHeader"]
            if FileHeader == "" :
                File['logbuf'] += "[ERROR] FileHeader is NOT Saved"
                return False
            
            MapFHeader = MapHWP.MapFileHeader(File, FileHeader)
            if MapFHeader["Property"] & 1 :
                File["Compression"] = True
            else :
                File["Compression"] = False
                
            if MapFHeader["Property"] & 2 :
                File["Password"] = True
                File['logbuf'] += "[ERROR] Password is required"
                return False
            else :
                File["Password"] = False
            
            
            # Trace Sector Data
            Record = ParseRecord()
            if File["ExtType"] != [] :                
                DecType = Record.DecryptDataRecord(File["ExtType"])
                if DecType == [] :
                    File['logbuf'] += "[ERROR] DecryptDataRecord( )"
                    return False
                File["DecType"] = DecType
            
            
            # Parse Decrypted Data Record
            if not Record.ParseDataRecord(File) :
                return False
            
            
            
        except :
            print traceback.format_exc()
            return False
        
        return True



class MappedHWP():
    def MapFileHeader(self, File, FileHeader):
        try :
            FHeader = {}
            OutFHeader = {}
            
            index = 0
            Position = 0
            for index in range( len(szFHeader) ) :
                if szFHeader[index] == 32 : 
                    Signature = BufferControl.Read(FileHeader, Position, szFHeader[index])
                    FHeader[ mFHeader[index] ] = BufferControl.ExtractAlphaNumber( Signature )
                elif szFHeader[index] == 4 :
                    FHeader[ mFHeader[index] ] = BufferControl.ReadDword(FileHeader, Position)
                else :
                    FHeader[ mFHeader[index] ] = binascii.b2a_hex( BufferControl.Read(FileHeader, Position, szFHeader[index]) )
                
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
#                        File['logbuf'] += DirEntry["EntryName"] + "\n                  "
                        fname = "%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"])
                        FileControl.WriteFile(fname, Sector)
                        
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
                File['logbuf'] += "[ERROR] RootEntry is NOT Saved"
                return False
            
            SSATable = File["SSATable"]
            RefSSAT = File["RefSSAT"]
            for RefEntry in RefSSAT :
                for DirEntry in DirList :
                    if RefEntry == "" :
                        continue                    
                        
                    if RefEntry == DirEntry["EntryName"] :
                        Sector = self.ExtractSectorbySSAT(pBuf, SSATable, DirEntry["SecID"], DirEntry["szData"])
#                        File['logbuf'] += DirEntry["EntryName"] + "\n                  "
                        fname = "%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"])
                        FileControl.WriteFile(fname, Sector)
                        
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
            
                tmpSector += BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x200)
                SecID = Table[SecID]
            
            Sector = BufferControl.Read(tmpSector, 0, Size)
            
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
            
                tmpSector += BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x40)
                SecID = Table[SecID]
            
            Sector = BufferControl.Read(tmpSector, 0, Size)
            
        except :
            print traceback.format_exc()
            return ""
        
        return Sector


class ParseRecord():
    def DecryptDataRecord(self, ExtType):
        try :
            DecType = []
            for fname in ExtType :            
                Buf = FileControl.ReadFileByBinary(fname)
                Buf_Dec = zlib.decompress(Buf, -15)
                sname = "%s_Decrypt" % fname
                FileControl.WriteFile(sname, Buf_Dec)
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
                
                File['logbuf'] += "\n\n\t    %s" % fname
                
                Position = 0
                DecrpytBuf = FileControl.ReadFileByBinary(fname)
                while Position < len( DecrpytBuf ) :
                    OverFlag = False
                    DataRecord = BufferControl.ReadDword(DecrpytBuf, Position)
                    
                    File['logbuf'] += "\n\t    0x%08X : " % Position
                    
                    Position += 4
                    TagID = DataRecord & 0x000003FF
                    Level = ( DataRecord & 0x000FFC00 ) >> 10
                    Size = ( DataRecord & 0xFFF00000 ) >> 20
                    if Size == 0xFFF :
                        OverFlag = True
                        Size = BufferControl.ReadDword(DecrpytBuf, Position)
                        Position += 4
                    
                    File['logbuf'] += "( 0x%08X   %03X   0x%08X (%s) ) - %s" % (TagID, Level, Size, OverFlag, HWPTAG[TagID])
                    
                    if OverFlag == True and Size > 0x00010000 :
                        File['logbuf'] += "   " + ">" * 10 + " Suspicious!!"
#                        BufferControl.PrintBuffer(File, DecrpytBuf, Position, 0x100)
                    
                    Position += Size
                    
        except :
            print traceback.format_exc()
            return False
        
        return True


    def ParseRecord(self, File, TagID, pBuf, Position, Size):
        try :
            File['logbuf'] += "\n\tParseRecord()"
        except :
            print traceback.format_exc()
            return False
        
        return True



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

#----------------------------------------------------------------------------------------------------------------------------------------------------#
#    DocInfo Structure 
#----------------------------------------------------------------------------------------------------------------------------------------------------#
# HWPTAG_DOCUMENT_PROPERTIES
szHWPTAG_DOCUMENT_PROPERTIES = [2,2,2,2,2,2,2,4,4,4,4]


# HWPTAG_ID_MAPPINGS
szHWPTAG_ID_MAPPINGS = [2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2]
# HWPTAG_ID_MAPPINGS - Value : 0        Binary Data
# HWPTAG_ID_MAPPINGS - Value : 1        Hangul Font
# HWPTAG_ID_MAPPINGS - Value : 2        English Font
# HWPTAG_ID_MAPPINGS - Value : 3        Hanja Font
# HWPTAG_ID_MAPPINGS - Value : 4        Japanese Font
# HWPTAG_ID_MAPPINGS - Value : 5        Etc Font
# HWPTAG_ID_MAPPINGS - Value : 6        Sign Font
# HWPTAG_ID_MAPPINGS - Value : 7        User Font
# HWPTAG_ID_MAPPINGS - Value : 8        Outline / Background
# HWPTAG_ID_MAPPINGS - Value : 9        Shape of letter
# HWPTAG_ID_MAPPINGS - Value : 10       Tab
# HWPTAG_ID_MAPPINGS - Value : 11       Paragraph Number
# HWPTAG_ID_MAPPINGS - Value : 12       Bullet 
# HWPTAG_ID_MAPPINGS - Value : 13       Shape of Paragraph 
# HWPTAG_ID_MAPPINGS - Value : 14       Style
# HWPTAG_ID_MAPPINGS - Value : 15       Shape of Memo

# HWPTAG_BIN_DATA


# HWPTAG_FACE_NAME


# HWPTAG_BORDER_FILL


# HWPTAG_CHAR_SHAPE


# HWPTAG_TAB_DEF


# HWPTAG_NUMBERING


# HWPTAG_BULLET


# HWPTAG_PARA_SHAPE


# HWPTAG_STYLE


# HWPTAG_DOC_DATA


# HWPTAG_DISTRIBUTE_DOC_DATA


# RESERVED


# HWPTAG_COMPATIBLE_DOCUMENT


# HWPTAG_LAYOUT_COMPATIBILITY


# HWPTAG_FORBIDDEN_CHAR

#----------------------------------------------------------------------------------------------------------------------------------------------------#
#    BodyText Structure 
#----------------------------------------------------------------------------------------------------------------------------------------------------#
# HWPTAG_PARA_HEADER


# HWPTAG_PARA_TEXT


# HWPTAG_PARA_CHAR


# HWPTAG_PARA_LINE_SEG


# HWPTAG_PARA_RANGE_TAG


# HWPTAG_CTRL_HEADER


# HWPTAG_LIST_HEADER


# HWPTAG_PAGE_DEF


# HWPTAG_FOOTNOTE_SHAPE


# HWPTAG_PAGE_BORDER_FILL


# HWPTAG_SHAPE_COMPONENT


# HWPTAG_TABLE


# HWPTAG_SHAPE_COMPONENT_LINE


# HWPTAG_SHAPE_COMPONENT_RECTANGLE


# HWPTAG_SHAPE_COMPONENT_ELLIPSE


# HWPTAG_SHAPE_COMPONENT_ARC


# HWPTAG_SHAPE_COMPONENT_POLYGON


# HWPTAG_SHAPE_COMPONENT_CURVE


# HWPTAG_SHAPE_COMPONENT_OLE


# HWPTAG_SHAPE_COMPONENT_PICTURE


# HWPTAG_SHAPE_COMPONENT_CONTAINER


# HWPTAG_CTRL_DATA


# HWPTAG_EQEDIT


# RESERVED


# HWPTAG_SHAPE_COMPONENT_TEXTART


# HWPTAG_FORM_OBJECT


# HWPTAG_SHAPE


# HWPTAG_MEMO_LIST


# HWPTAG_CHART_DATA


# HWPTAG_SHAPE_COMPONENT_UNKNOWN

