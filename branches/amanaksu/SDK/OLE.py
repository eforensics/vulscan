# External Import
from struct import error, unpack
from traceback import format_exc
from collections import namedtuple
from string import letters, digits

# Internal Import

class COLE():
    @classmethod
    def fnIsOLE(cls, s_buffer):
        
        try :
            
            if (s_buffer[0:4] == "\xd0\xcf\x11\xe0" and s_buffer[4:8] == "\xa1\xb1\x1a\xe1") or (s_buffer[0:4] == "\xd0\xcd\x11\xe0") : 
                return "OLE"
            else :
                return None
            
        except error : 
            return None
        
        except :
            print format_exc()
            return None
    @classmethod
    def fnIsHWP23(cls, s_buffer):
        
        try :
            
            s_HWP23Sig = s_buffer[:0x20]
            if (s_HWP23Sig.find("HWP Document File V2.00") != -1) or (s_HWP23Sig.find("HWP Document File V3.00") != -1) :
                return "HWP23"
            else :
                return None
            
        except :
            print format_exc()
            return None
        
class CStructOLE():
    def fnStructOLEHeader(self, s_buffer):
        t_OLEHeader = ()        
        
        try :
            
            MappedOLE = CMappedOLE()
            t_OLEHeader = MappedOLE.fnMapOLEHeader(s_buffer[0:SIZE_OF_OLE_HEADER])
            if t_OLEHeader == () :
                print "[-] Error - fnMapOLEHeader()"
        
        except :
            print format_exc()
        
        return t_OLEHeader
    def fnStructOLEMSAT(self, s_buffer, s_MSAT, n_MSATSecID, n_NumMSAT):
        t_MSAT = ()
        
        try :
            # Check Exception
            if n_NumMSAT < 0 :
                print "[-] Error - MSAT Number : %d" % n_NumMSAT
            
            # MSAT List
            else :
                t_MSAT = unpack("109l", s_MSAT)
                
                # Extra MSAT
                if n_NumMSAT != 0 :
                    n_ExtraMSATCnt = n_NumMSAT
                    n_ExtraMSATSecID = n_MSATSecID
                    t_tmpMSAT = ()
                    
                    while True :
                        n_ExtraMSATSecPos = (n_ExtraMSATSecID + 1) * SIZE_OF_SECTOR
                        s_tmpMSAT = s_buffer[n_ExtraMSATSecPos:n_ExtraMSATSecPos + SIZE_OF_SECTOR]
                        t_tmpMSAT = unpack("128l", s_tmpMSAT)
                        t_MSAT += t_tmpMSAT[:-1]
                        
                        # First Index is "0", Last Index is "MaxIndex -1"
                        n_ExtraMSATSecID = t_tmpMSAT[128-1] 
                        
                        n_ExtraMSATCnt -= 1
                        if n_ExtraMSATCnt == 0 :
                            break
            
        except :
            print format_exc()
        
        return t_MSAT
    def fnStructOLESAT(self, s_buffer, t_MSAT):
        t_SAT = ()
        
        try :
            
            for n_SecID in t_MSAT :
                if n_SecID < 0 :
                    continue
                
                n_SecPos = (n_SecID + 1) * SIZE_OF_SECTOR
                s_tmpSAT = s_buffer[n_SecPos:n_SecPos + SIZE_OF_SECTOR]
                t_tmpSAT = unpack( "128l", s_tmpSAT )
                t_SAT += t_tmpSAT
            
        except :
            print format_exc()
            
        return t_SAT        
    def fnStructOLESSAT(self, s_buffer, t_SAT, n_SSATSecID, SIZE_OF_SECTOR):
        t_SSAT = ()
        
        try :
            
            s_SSAT = self.fnStructOLESector(s_buffer, t_SAT, n_SSATSecID, SIZE_OF_SECTOR)
            if s_SSAT == "" :
                return None
            
            t_SSAT = unpack( str(s_SSAT.__len__() / 4) + "l", s_SSAT )
            
        except :
            print format_exc()
        
        return t_SSAT
    def fnStructOLESector(self, s_buffer, t_Table, n_SecID, n_Size):
        s_Sector = ""
        
        try :
            ExceptOLE = CExceptOLE()
            
            if n_Size == SIZE_OF_SECTOR :
                n_AddSector = 1
            if n_Size == SIZE_OF_SHORT_SECTOR :
                n_AddSector = 0
                
            while True :
                s_tmpSector = ""
                n_SecPos = (n_SecID + n_AddSector) * n_Size
                s_tmpSector = s_buffer[n_SecPos:n_SecPos + n_Size]
                if s_tmpSector == "" :
                    if not ExceptOLE.fnExceptTable() :
                        return ""
                
                s_Sector += s_tmpSector
                n_SecID = t_Table[ n_SecID ]
                if n_SecID <= 0 :
                    break 
                
        except :
            print format_exc()
            
        return s_Sector
    def fnStructOLEDirectory(self, s_Directory):
        l_Directory = []
        
        try :
            MappedOLE = CMappedOLE()
            
            if (s_Directory.__len__() % SIZE_OF_DIRECTORY) != 0 :
                print "[-] Error - Directory Buffer Size ( %d )" % s_Directory.__len__()
            
            else :
                n_Start = 0
                while True :
                    s_SubDirectory = s_Directory[n_Start:n_Start + SIZE_OF_DIRECTORY]
                    t_Directory = MappedOLE.fnMapOLEDirectory(s_SubDirectory)
                    l_Directory.append( list(t_Directory) )
                    n_Start += SIZE_OF_DIRECTORY
                    if n_Start == s_Directory.__len__() :
                        break
            
        except :
            print format_exc()
            
        return l_Directory
        
class CMappedOLE(CStructOLE):
    def fnMapOLEHeader(self, s_buffer):
        t_OLEHeader = ()
        
        try :
            
            t_OLEHeader_Name = namedtuple("OLEHeader", RULE_OLEHEADER_NAME)
            t_OLEHeader = t_OLEHeader_Name._make( unpack(RULE_OLEHEADER_PATTERN, s_buffer) )
            
        except :
            print format_exc()
        
        return t_OLEHeader
    def fnMapOLEDirectory(self, s_Directory):
        t_OLEDirectory = ()
        
        try :
            
            t_OLEDirectory_Name = namedtuple("OLEDirectory", RULE_OLEDIRECTORY_NAME)
            t_OLEDirectory = t_OLEDirectory_Name._make( unpack(RULE_OLEDIRECTORY_PATTERN, s_Directory) )
            
        except :
            print format_exc()
            
        return t_OLEDirectory

class CExceptOLE(CStructOLE):
    def fnExceptTable(self, s_buffer, n_SecID, n_SecPos, n_Size):
        
        try :
            
            if n_SecPos >= s_buffer.__len__() or n_SecPos + n_Size >= s_buffer.__len__() :
                print "[-] Error - Out of range()"
                print "\t" + "- buffer size : 0x%08X, SecID : 0x%04X, Complex Value : 0x%04X, Position : 0x%08X" % (s_buffer.__len__(), n_SecID, n_Size, n_SecPos)
            else :
                print "[-] Error - Unknown ( SecID : 0x%08X, Position : 0x%08X )" % (n_SecID, n_SecPos)
            
        except :
            print format_exc()
            return False
        
        return True
    
# Define

SIZE_OF_SECTOR = 0x200
SIZE_OF_SHORT_SECTOR = 0x40

SIZE_OF_OLE_HEADER = 0x200
RULE_OLEHEADER_NAME = ('Signature UID Revision Version ByteOrder szSector szShort Reserved1 NumSAT DirSecID Reserved2 szMinSector SSATSecID NumSSAT MSATSecID NumMSAT MSAT')
RULE_OLEHEADER_PATTERN = '=8s16s5h10s8l436s'

SIZE_OF_DIRECTORY = 0x80
RULE_OLEDIRECTORY_NAME = ('DirName Size Type Color LeftChild RightChild RootChild UID Flags createTime LastModifyTime SecID_Start StreamSize Reserved')
RULE_OLEDIRECTORY_PATTERN = '=64sH2B3l16sl8s8s3l'

class COffice():
    pass

class CStructHWP():
    def fnStructHWPHeader(self, s_buffer, dl_OLEDirectory, t_SAT, t_SSAT):
        t_HWPHeader = ()
        
        try :
            s_RootEntry = self.fnStructHWPSector(s_buffer, None, dl_OLEDirectory, t_SAT, t_SSAT, "RootEntry")
            if s_RootEntry == "" :
                print "[-] Error - StructHWPHeader( RootEntry )"
                return t_HWPHeader
            
            s_HWPHeader = self.fnStructHWPSector(s_buffer, s_RootEntry, dl_OLEDirectory, t_SAT, t_SSAT, "FileHeader")
            if s_HWPHeader == "" :
                print "[-] Error - StructHWPHeader( FileHeader )"
                return t_HWPHeader
            
            MappedHWP = CMappedHWP()
            t_HWPHeader = MappedHWP.fnMapHWPHeader(s_HWPHeader)
            if t_HWPHeader == () :
                print "[-] Error - MapHWPHeader()"
                return t_HWPHeader
            
            if t_HWPHeader.Property & 2 :
                print "[-] Password is required"
                return t_HWPHeader
            
        except :
            print format_exc()
        
        return t_HWPHeader
    def fnStructHWPSector(self, s_buffer, s_RootEntry, dl_OLEDirectory, t_SAT, t_SSAT, s_Entry ):
        s_Sector = ""
        
        try :
            UtilOLE = CUtilOLE()
            n_Index = UtilOLE.fnFindEntryIndex(dl_OLEDirectory, s_Entry)
            if n_Index == None :
                print "[-] Error - FindEntryIndex( %s )" % s_Entry
                return s_Sector
            
            
            n_EntryType = dl_OLEDirectory[n_Index][2]
            n_EntrySecID = dl_OLEDirectory[n_Index][11]
            n_EntrySize = dl_OLEDirectory[n_Index][12]
            
            if n_EntryType == 5 or n_EntrySize >= 0x1000 :
                s_EntryBuffer = s_buffer
                t_Table = t_SAT
                n_Size = SIZE_OF_SECTOR
            elif n_EntryType != 5 and n_EntrySize < 0x1000 :
                if s_RootEntry == None :
                    print "[-] Error - StructHWPSector()'s RootEntry Parameter ( %s - 0x%08X )" % (s_Entry, n_EntryType)
                    return s_Sector
                
                s_EntryBuffer = s_RootEntry
                t_Table = t_SSAT
                n_Size = SIZE_OF_SHORT_SECTOR
            else :
                print "[-] Error - Do not Referred Type"
                return s_Sector
            
            StructOLE = CStructOLE()
            s_Sector = StructOLE.fnStructOLESector(s_EntryBuffer, t_Table, n_EntrySecID, n_Size)
            if s_Sector == "" :
                print "[-] Error - StructOLESector() for %s in HWP" % s_Entry
            
        except :
            print format_exc()
        
        return s_Sector[:n_EntrySize]
    def fnStructHWPShortSector(self, s_buffer, dl_OLEDirectory, t_SSAT, b_Flag):
        l_SSector = []
        l_SSector_Name = []
        l_SSector_Data = []
        
        try :
            StructOLE = CStructOLE()
            UtilOLE = CUtilOLE()
            
            n_Cnt = 0
            
            while n_Cnt < dl_OLEDirectory.__len__() :
                if dl_OLEDirectory[n_Cnt][2] != 5 and dl_OLEDirectory[n_Cnt][12] < 0x1000 :
                    s_SSector = StructOLE.fnStructOLESector(s_buffer, t_SSAT, dl_OLEDirectory[n_Cnt][11], SIZE_OF_SHORT_SECTOR)
                    if s_SSector == "" :
                        print "[-] Error - StructHWPShortSector() for %s in HWP" % (UtilOLE.fnExtractAlphaNumber(dl_OLEDirectory[n_Cnt][0]))
                        return l_SSector
            
                    l_SSector_Name.append( UtilOLE.fnExtractAlphaNumber(dl_OLEDirectory[n_Cnt][0]) )
                    l_SSector_Data.append( s_SSector )
            
                n_Cnt += 1
            
            l_SSector.append( l_SSector_Name )
            l_SSector.append( l_SSector_Data )
            
        except :
            print format_exc()
            
        return l_SSector

class CMappedHWP(CStructHWP):
    def fnMapHWPHeader(self, s_buffer):
        t_HWPHeader = ()
        
        try :
            
            t_HWPHeader_Name = namedtuple("HWPHeader", RULE_HWPHEADER_NAME)
            t_HWPHeader = t_HWPHeader_Name._make( unpack(RULE_HWPHEADER_PATTERN, s_buffer) )
            
        except :
            print format_exc()
            
        return t_HWPHeader
    
RULE_HWPHEADER_NAME = ('Signature Version Property Reserved')
RULE_HWPHEADER_PATTERN = '=32s2l216s'

class CUtilOLE():
    def fnExtractAlphaNumber(self, s_buffer):
        s_Outbuffer = ""
        
        try :
            
            s_Pattern = letters + digits
            l_ExtList = []
            for s_OneByte in s_buffer :
                if s_OneByte in s_Pattern :
                    l_ExtList.append( s_OneByte )
            
            s_Outbuffer = "".join( l_ExtList )
            
        except :
            print format_exc()
            
        return s_Outbuffer
    def fnFindEntryIndex(self, dl_OLEDirectory, s_Entry):
        n_RetIndex = None
        
        try :
            
            n_Index = 0
            while n_Index < dl_OLEDirectory.__len__() :
                s_Repaired = self.fnExtractAlphaNumber(dl_OLEDirectory[n_Index][0])
                if s_Repaired[:s_Entry.__len__()].find( s_Entry ) != -1 :
                    n_RetIndex = n_Index
                    break
                n_Index += 1
            
        except :
            print format_exc()
            
        return n_RetIndex

class CPrintOLE():
    def fnPrintOLEHeader(self):
        pass
    def fnPrintOLEDirectory(self):
        pass

class CPrintHWP():
    pass

class CPrintOffice():
    pass




