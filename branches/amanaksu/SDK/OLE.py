# External Import
from struct import error, unpack
from traceback import format_exc
from collections import namedtuple
from string import letters, digits

# Internal Import
try :
    from Common import CBuffer
except :
    print "[-] Error - Internal Import "
    exit(-1)

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
    def fnFindDirectory(self, s_buffer):
        dl_OLEDirectory = []
        
        try :
            
            StructOLE = CStructOLE()
            
            # Parse OLE Header
            t_OLEHeader = StructOLE.fnStructOLEHeader(s_buffer)
            
            # Extract OLE MSAT List
            t_MSAT = StructOLE.fnStructOLEMSAT(s_buffer, t_OLEHeader)
            if t_MSAT == () :
                print "[-] Error - fnStructOLEMSAT()"
                return dl_OLEDirectory
            
            # Extract OLE SAT List
            t_SAT = StructOLE.fnStructOLESAT(s_buffer, t_MSAT)
            if t_SAT == () :
                print "[-] Error - fnStructOLESAT()"
                return dl_OLEDirectory
            
            # Parse Directory Sector
            dl_OLEDirectory = StructOLE.fnStructOLEDirectoryEx(s_buffer, t_SAT, t_OLEHeader.DirSecID, SIZE_OF_SECTOR)
            if dl_OLEDirectory == [] :
                print "[-] Error - fnStructOLEDirectoryEx()"
                return dl_OLEDirectory
            
        except :
            print format_exc()
            
        return dl_OLEDirectory
    def fnIsSubFormat(self, s_buffer ):
        s_format = ""
        
        try :
            
            dl_OLEDirectory = self.fnFindDirectory(s_buffer)
            if dl_OLEDirectory == [] :
                print "[-] Error - fnFindDirectory()"
                return s_format
            
            UtilOLE = CUtilOLE()
            n_Index = UtilOLE.fnFindEntryIndex(dl_OLEDirectory, "Hwp")
            if n_Index == None :
                s_format = "Office"
            if n_Index != None :
                s_format = "HWP"
            
        except :
            print format_exc()
        
        return s_format
       
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
    def fnStructOLEMSAT(self, s_buffer, t_OLEHeader):
        t_MSAT = ()
        
        try :
            ExceptOLE = CExceptOLE()
            
            # Init
            s_MSAT = t_OLEHeader.MSAT
            n_MSATSecID = t_OLEHeader.MSATSecID
            n_NumMSAT = t_OLEHeader.NumMSAT
            
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
            
            # Check MSAT
            if not ExceptOLE.fnExceptMSAT(t_MSAT, t_OLEHeader.NumSAT) :
                return ()
            
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
    def fnStructOLESectorEx(self, s_buffer1, s_buffer2, dl_OLEDirectory, t_SAT, t_SSAT, s_Entry):
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
                s_EntryBuffer = s_buffer1
                t_Table = t_SAT
                n_Size = SIZE_OF_SECTOR
            elif n_EntryType != 5 and n_EntrySize < 0x1000 :
                if s_buffer2 == None :
                    print "[-] Error - StructHWPSector()'s RootEntry Parameter ( %s - 0x%08X )" % (s_Entry, n_EntryType)
                    return s_Sector
                
                s_EntryBuffer = s_buffer2
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
    def fnStructOLEDirectoryEx(self, s_buffer, t_SAT, n_DirSecID, n_Size):
        dl_OLEDirectory = []
        
        try :
            
            s_Directory = self.fnStructOLESector(s_buffer, t_SAT, n_DirSecID, n_Size)
            if s_Directory == "" :
                print "[-] Error - fnStructOLESector()"
                return dl_OLEDirectory
            
            dl_OLEDirectory = self.fnStructOLEDirectory(s_Directory)
            if dl_OLEDirectory == [] :
                print "[-] Error - fnStructOLEDirectory()"
                return dl_OLEDirectory
            
        except :
            print format_exc()
            
        return dl_OLEDirectory
 
class CMappedOLE(CStructOLE):
    t_OLEHeader = ()
    t_OLEDirectory = ()
    
    def fnMapOLEHeader(self, s_buffer):        
        
        try :
            
            t_OLEHeader_Name = namedtuple("OLEHeader", RULE_OLEHEADER_NAME)
            self.t_OLEHeader = t_OLEHeader_Name._make( unpack(RULE_OLEHEADER_PATTERN, s_buffer) )
            
        except :
            print format_exc()
        
        return self.t_OLEHeader
    def fnMapOLEDirectory(self, s_Directory):
        
        try :
            
            t_OLEDirectory_Name = namedtuple("OLEDirectory", RULE_OLEDIRECTORY_NAME)
            self.t_OLEDirectory = t_OLEDirectory_Name._make( unpack(RULE_OLEDIRECTORY_PATTERN, s_Directory) )
            
        except :
            print format_exc()
            
        return self.t_OLEDirectory

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
    def fnExceptMSAT(self, t_MSAT, n_NumSAT):
        
        try :
            
            n_SATCnt = 0
            for n_SATIndex in t_MSAT :
                if n_SATIndex >= 0 :
                    n_SATCnt += 1
            
            if n_SATCnt != n_NumSAT :
                print "[-] Error - SAT Count ( Checked Count : 0x%08X, NumSAT : 0x%08X )" % (n_SATCnt, n_NumSAT)
                return False 
            
        except :
            print format_exc()
            
        return True

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

class CStructOffice():
    def fnStructOfficeHeader(self, s_buffer, dl_OLEDirectory, t_SAT, t_SSAT):
        t_WordDoc = ()
        
        try :
            StructOLE = CStructOLE()
            s_WordDoc = StructOLE.fnStructOLESectorEx(s_buffer, None, dl_OLEDirectory, t_SAT, t_SSAT, "WordDocument")
            if s_WordDoc == "" :
                print "[-] Error - StructOfficeHeader( WordDocument )"
                return t_WordDoc
            
            if not self.fnStructOfficeWordDoc(s_WordDoc) :
                print "[-] Error - StructofficeWordDoc()"
                return t_WordDoc
            
        except :
            print format_exc()
        
        return t_WordDoc
    def fnStructOfficeWordDoc(self, s_WordDoc):
        
        try :
            n_Position = 0
            
            # Mapped FibBase
            n_Position = self.fnStructFibBase(s_WordDoc, n_Position)
            if n_Position == None :
                print "[-] Error - fnStructFibBase()"
                return False
            
            # Mapped FibRgW97
            n_Position = self.fnStructGetSizeFibRgW97(s_WordDoc, n_Position)
            if n_Position == None :
                print "[-] Error - fnStructGetSizeFibW97()"
                return False
            
            n_Position = self.fnStructFibRgW97(s_WordDoc, n_Position)
            if n_Position == None :
                print "[-] Error - fnStructFibW97()"
                return False
            
            # Mapped FibRgLw97
            n_Position = self.fnStructGetSizeFibRgLw97(s_WordDoc, n_Position)
            if n_Position == None :
                print "[-] Error - fnStructGetSizeFibLw97()"
                return False
            
            n_Position = self.fnStructFibRgLw97(s_WordDoc, n_Position)
            if n_Position == None :
                print "[-] Error - fnStructFibLw97()"
                return False
            
            # Mapped cbRgFcLcb 
            n_Position, n_Position_FibRgFcLcb = self.fnStructGetSizecbRgFcLcb(s_WordDoc, n_Position)
            
            
            # Fixed nFib
            n_Fib = self.fnStructSetFib(s_WordDoc, n_Position)
            
            # Mapped FibRgFcLcbBlob
            if not self.fnStructFibRgFcLcbBlob(s_WordDoc, n_Position_FibRgFcLcb, n_Fib) :
                print "[-] Error - fnStructFibRgFcLcbBlob()"
                return False
            
        except :
            print format_exc()
            
        return True
    def fnStructFibBase(self, s_buffer, n_Position):
        
        try :
            MappedOffice = CMappedOffice()
            ExceptOffice = CExceptOffice()
            
            MappedOffice.fnMapFibBase(s_buffer[n_Position:n_Position + SIZE_OF_FIBBASE])
            if not ExceptOffice.fnExceptFibBase( MappedOffice.t_FibBase ) :
                return None
            
            n_Position += SIZE_OF_FIBBASE
            
        except :
            print format_exc()
        
        return n_Position
    def fnStructGetSizeFibRgW97(self, s_buffer, n_Position):
        
        try :
            
            n_CSW = CBuffer.fnReadData(s_buffer, n_Position, 2)
            if n_CSW != 0x000E :
                print "[-] Error ( Position : 0x%08X, CSW : 0x%04X )" % (n_Position, n_CSW)
                return None
            
            n_Position += 2
            n_szFibRgW97 = n_CSW * 2
            if n_szFibRgW97 != SIZE_OF_FIBW97 :
                print "[-] Error ( Position : 0x%08X, szFibW97 : 0x%04X ) " % (n_Position, n_szFibRgW97) 
                return None
            
        except :
            print format_exc()
            
        return n_Position
    def fnStructFibRgW97(self, s_buffer, n_Position):
        
        try :
            MappedOffice = CMappedOffice()
            ExceptOffice = CExceptOffice()
            
            MappedOffice.fnMapFibRgW97(s_buffer[n_Position:n_Position + SIZE_OF_FIBW97])
            if not ExceptOffice.fnExceptFibRgW97( MappedOffice.t_FibW97 ) :
                return None
            
            n_Position += SIZE_OF_FIBW97
            
        except :
            print format_exc()
            
        return n_Position
    def fnStructGetSizeFibRgLw97(self, s_buffer, n_Position):
        
        try :
            
            n_CSLW = CBuffer.fnReadData(s_buffer, n_Position, 2)
            if n_CSLW != 0x0016 :
                print "[-] Error ( Position : 0x%08X, CSLW : 0x%04X )" % (n_Position, n_CSLW)
                return None
            
            n_Position += 2
            n_szFibRgLw97 = n_CSLW * 4
            if n_szFibRgLw97 != SIZE_OF_FIBLW97 :
                print "[-] Error ( Position : 0x%08X, szFibLw97 : 0x%04X )" % (n_Position, n_szFibRgLw97)
                return None
            
        except :
            print format_exc()
            
        return n_Position
    def fnStructFibRgLw97(self, s_buffer, n_Position):
        
        try :
            MappedOffice = CMappedOffice()
            ExceptOffice = CExceptOffice()
            
            MappedOffice.fnMapFibRgLw97(s_buffer[n_Position:n_Position + SIZE_OF_FIBLW97])
            if not ExceptOffice.fnExceptFibRgLw97( MappedOffice.t_FibLw97 ) :
                return None
            
            n_Position += SIZE_OF_FIBLW97
            
        except :
            print format_exc()
            
        return n_Position
    def fnStructGetSizecbRgFcLcb(self, s_buffer, n_Position):
        
        try :
            l_cbRgFcLcb = [0x005D, 0x006C, 0x0088, 0x00A4, 0x00B7]
            
            n_cbRgFcLcb = CBuffer.fnReadData(s_buffer, n_Position, 2)
            if not n_cbRgFcLcb in l_cbRgFcLcb :
                print "[-] Error ( Position : 0x%08X, cbRgFcLcb : 0x%04X )" % (n_Position, n_cbRgFcLcb)
                return None
            
            n_Position += 2
            n_OldPosition = n_Position 
            
            n_szRgFcLcb = n_cbRgFcLcb * 8
            
            n_Position += n_szRgFcLcb
            
        except :
            print format_exc()
            
        return n_Position, n_OldPosition
    def fnStructSetFib(self, s_buffer, n_Position) :
        
        try :
            MappedOffice = CMappedOffice()
            
            l_cswNew = [0x0000, 0x0002, 0x0002, 0x0002, 0x0005]
            l_Fib = [0x00C1, 0x00D9, 0x0101, 0x010C, 0x0112]
            
            n_cswNew = CBuffer.fnReadData(s_buffer, n_Position, 2)
            if not n_cswNew in l_cswNew :
                print "[-] Error ( Position : 0x%08X, cswNew : 0x%04X )" % (n_Position, n_cswNew)
                return None
            
            n_Position += 2
            
            if n_cswNew == 0 :
                n_Fib = MappedOffice.t_FibBase.nFib
            else :
                n_Fib = CBuffer.fnReadData(s_buffer, n_Position, 2)
            
            if not n_Fib in l_Fib :
                print "[-] Error ( cswNew : 0x%04X, nFib : 0x%04X )" % (n_cswNew, n_Fib)
                return None
            
        except :
            print format_exc()
            
        return n_Fib
    def fnStructFibRgFcLcbBlob(self, s_buffer, n_Position, n_Fib):
        
        try :
            MappedOffice = CMappedOffice()
            ExceptOffice = CExceptOffice()
            
            MappedOffice.fnMapFibRgFcLcbBlob(s_buffer[n_Position:], n_Fib)
            if not ExceptOffice.fnExceptFibRgFcLcbBlob( MappedOffice.l_FibRgFcLcbBlob ) :
                return False
            
        except :
            print format_exc()
            return False
            
        return True
            
class CMappedOffice(CStructOffice): 
    t_FibBase = ()
    t_FibW97 = ()
    t_FibLw97 = ()
    l_FibRgFcLcbBlob = []
    t_FibRgFcLcb97 = ()
    t_FibRgFcLcb2000 = ()
    t_FibRgFcLcb2002 = ()
    t_FibRgFcLcb2003 = ()
    t_FibRgFcLcb2007 = ()
    
    def fnMapFibBase(self, s_FibBase):
        
        try :
        
            t_FibBase_Name = namedtuple("FibBase", RULE_FIBBASE_NAME)
            self.t_FibBase = t_FibBase_Name._make( unpack(RULE_FIBBASE_PATTERN, s_FibBase) )
            
        except :
            print format_exc()
        
        return self.t_FibBase
    def fnMapFibRgW97(self, s_FibRgW97):
        
        try :
            
            t_FibW97_Name = namedtuple("FibW97", RULE_FIBW97_NAME)
            self.t_FibW97 = t_FibW97_Name._make( unpack(RULE_FIBW97_PATTERN, s_FibRgW97) )
            
        except :
            print format_exc()
            
        return self.t_FibW97
    def fnMapFibRgLw97(self, s_FibRgLw97):
        
        try :
            
            t_FibLw97_Name = namedtuple("FibLw97", RULE_FIBLW97_NAME)
            self.t_FibLw97 = t_FibLw97_Name._make( unpack(RULE_FIBLW97_PATTERN, s_FibRgLw97) )
            
        except :
            print format_exc()
            
        return self.t_FibLw97
    def fnMapFibRgFcLcbBlob(self, s_FibRgFcLcb, n_Fib):
        
        try :
            l_Blob = ["97",   "2000", "2002", "2003", "2007"]
            l_Size = [0x2E8,  0x78,   0xE0,   0xE0,   0x98  ]
            l_Fib =  [0x00C1, 0x00D9, 0x0101, 0x010C, 0x0112]
            
            n_Position = 0
            
            for n_VerYear in enumerate(l_Blob) :
                eval( "self.fnMapFibRgFcLcb%s(s_FibRgFcLcb[n_Position:n_Position + SIZE_OF_FIBFCLCB%s])" % (n_VerYear[1], l_Size[ n_VerYear[0] ]) )
                if eval( "self.t_FibRgFcLcb%s" % n_VerYear[1] ) == () :
                    print "[-] Error - fnMapFibRgFcLcb%s( Position : 0x%08X, Size : 0x%08X )" % (n_VerYear[1], n_Position, l_Size[ n_VerYear[0] ] )
                    break
                
                self.l_FibRgFcLcbBlob.append( eval( "self.t_FibRgFcLcb%s" % n_VerYear[1] ) )
                if l_Fib[ n_VerYear[0] ] == 0x0112 :
                    if n_Fib != 0x0112 :
                        print "[-] Error - nFib : 0x%04X" % n_Fib
                    break
                elif n_Fib == l_Fib[ n_VerYear[0] ] :
                    break
            
                n_Position += l_Size[ n_VerYear[0] ]

        except :
            print format_exc()
            
        return self.l_FibRgFcLcbBlob
    def fnMapFibRgFcLcb97(self, s_FibRgFcLcb97):
        
        try :
            
            t_FibRgFcLcb97_Name = namedtuple("FibRgFcLcb97", RULE_FIBFCLCB_NAME_97)
            self.t_FibRgFcLcb97 = t_FibRgFcLcb97_Name._make( unpack(RULE_FIBFCLCB_PATTERN_97, s_FibRgFcLcb97) )
            
        except :
            print format_exc()
            
        return self.t_FibRgFcLcb97
    def fnMapFibRgFcLcb2000(self, s_FibRgFcLcb2000):
        
        try :
            
            t_FibRgFcLcb2000_Name = namedtuple("FibRgFcLcb2000", RULE_FIBFCLCB_NAME_2000)
            self.t_FibRgFcLcb2000 = t_FibRgFcLcb2000_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2000, s_FibRgFcLcb2000) )
            
        except :
            print format_exc()
            
        return self.t_FibRgFcLcb2000
    def fnMapFibRgFcLcb2002(self, s_FibRgFcLcb2002):
        
        try :
            
            t_FibRgFcLcb2002_Name = namedtuple("FibRgFcLcb2002", RULE_FIBFCLCB_NAME_2002)
            self.t_FibRgFcLcb2002 = t_FibRgFcLcb2002_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2002, s_FibRgFcLcb2002) )
            
        except :
            print format_exc()
            
        return self.t_FibRgFcLcb2002
    def fnMapFibRgFcLcb2003(self, s_FibRgFcLcb2003):
        
        try :
            
            t_FibRgFcLcb2003_Name = namedtuple("FibRgFcLcb2003", RULE_FIBFCLCB_NAME_2003)
            self.t_FibRgFcLcb2003 = t_FibRgFcLcb2003_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2003, s_FibRgFcLcb2003) )
            
        except :
            print format_exc()
            
        return self.t_FibRgFcLcb2003
    def fnMapFibRgFcLcb2007(self, s_FibRgFcLcb2007):
        
        try :
            
            t_FibRgFcLcb2007_Name = namedtuple("FibRgFcLcb2007", RULE_FIBFCLCB_NAME_2007)
            self.t_FibRgFcLcb2007 = t_FibRgFcLcb2007_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2007, s_FibRgFcLcb2007) )
            
        except :
            print format_exc()
            
        return self.t_FibRgFcLcb2007

class CStructStream():
    pass

class CMappedStream():
    pass

class CUtilOffice(CStructOffice):
    pass

class CExceptOffice(CStructOffice):
    def fnExceptFibBase(self, t_FibBase):
        
        try :
            l_FIBMagicNumber            = [0xA5DC,                0xA5EC         ]
            l_VersionByFIBMagicNumber   = ["Word 6.0 / 7.0 (95)", "Word 8.0 (97)"]
            
            if t_FibBase == () :
                print "[-] Error - fnMapFibBase()"
                return False
            
            if t_FibBase.wIdent not in l_FIBMagicNumber :
                print "[-] Error - FibBase's Signature( 0x%08X )" % t_FibBase.wIdent
                return False
            
            for n_Ver in enumerate(l_FIBMagicNumber) :
                if t_FibBase.wIdent == n_Ver[1] :
                    print "[*] Version By Fib MagicNumber : %s" % l_VersionByFIBMagicNumber[ n_Ver[0] ]
            
        except :
            print format_exc()
            
        return True
    def fnExceptFibRgW97(self, t_FibRgW97):
        
        try :
            
            if t_FibRgW97 == () :
                print "[-] Error - fnMapFibRgW97()"
                return False
            
        except :
            print format_exc()
            
        return True
    def fnExceptFibRgLw97(self, t_FibRgLw97):
        
        try :
            
            if t_FibRgLw97 == () :
                print "[-] Error - fnMapFibRgLw97()"
                return False
        
        except :
            print format_exc()
            
        return True
    def fnExceptFibRgFcLcbBlob(self, l_FibRgFcLcb):
        
        try :
            
            if l_FibRgFcLcb == [] :
                print "[-] Error -fnMapFibRgFcLcbBlob()"
                return False
            
        except :
            print format_exc()
        
        return True

SIZE_OF_FIBBASE = 0x20            # 32Bytes
RULE_FIBBASE_NAME = ('wIdent nFib unused lid pnNext Flag nFibBack lKey envr envtFlag reserved3 reserved4 reserved5 reserved6')
RULE_FIBBASE_PATTERN = '=7H1L2B2H2L'

SIZE_OF_FIBW97 = 0x1C             # 28Bytes
RULE_FIBW97_NAME = ('reserved1 reserved2 reserved3 reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 lidFE')
RULE_FIBW97_PATTERN = '=14H'

SIZE_OF_FIBLW97 = 0x58            # 88Bytes
RULE_FIBLW97_NAME = ('cbMac reserved1 reserved2 ccpText ccpFtn ccpHdd reserved3 ccpAtn ccpEdn ccpTxbx ccpHdrTxbx reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 reserved14')
RULE_FIBLW97_PATTERN = '=22L'

SIZE_OF_FIBFCLCB_97 = 0x2E8       # 744Bytes
RULE_FIBFCLCB_NAME_97 = ('fcStshfOrig lcbStshfOrig fcStshf lcbStshf fcPlcffndRef lcbPlcffndRef fcPlcffndTxt lcbPlcffndTxt fcPlcfandRef lcbPlcfandRef fcPlcfandTxt lcbPlcfandTxt fcPlcfSed lcbPlcfSed fcPlcPad lcbPlcPad fcPlcfPhe lcbPlcfPhe fcSttbfGlsy lcbSttbfGlsy fcPlcfGlsy lcbPlcfGlsy fcPlcfHdd lcbPlcfHdd fcPlcBteChpx lcbPlcBteChpx fcPlcfBtePapx lcbPlcfBtePapx fcPlcfSea lcbPlcfSea fcSttbfFfn lcbSttbfFfn fcPlcfFldMom lcbPlcfFldMom fcPlcfFldHdr lcbPlcfFldHdr fcPlcfFldFtn lcbPlcfFldFtn fcPlcfFldAtn lcbPlcfFldAtn fcPlcfFldMcr lcbPlcfFldMcr fcSttbfBkmk lcbSttbfBkmk fcPlcfBkf lcbPlcfBkf fcPlcfBkl lcbPlcfBkl fcCmds lcbCmds fcUnused1 lcbUnused1 fcSttbfMcr lcbSttbfMcr fcPrDrvr lcbPDrvr fcPrEnvPort lcbPrEnvPort fcPrEnvLand lcbPrEnvLand fcWss lcbWss fcDop lcbDop fcSttbfAssoc lcbSttbfAssoc fcClx lcbClx fcPlcfPgdFtn lcbPlcfPgdFtn fcAutosaveSource lcbAutosaveSource fcGrpXstAtnOwners lcbGrpXstAtnOwners fcSttbfAtnBkmk lcbSttbfAtnBkmk fcUnused3 lcbUnused3 fcUnused2 lcbUnused2 fcPlcSpaMon lcbPlcSpaMon fcPlcSpaHdr lcbPlcSpaHdr fcPlcAtnBkf lcbPlcAtnBkf fcPlcAtnBkl lcbPlcAtnBkl fcPms lcbPms fcFormFldSttbs lcbFormFldSttbs fcPlcfendRef lcbPlcfendRef fcPlcfendTxt lcbPlcfendTxt fcPlcfFldEdn lcbPlcfFldEdn fcUnused4 lcbUnused4 fcDggInfo lcbDggInfo fcSttbfRMark lcbSttbfRMark fcSttbfCaption lcbSttbfCaption fcSttbfAutoCaption lcbSttbfAutoCaption fcPlcWkb lcbPlcWkb fcPlcfSpl lcbPlcfSpl fcPlcftxbxTxt lcbPlcftxbxTxt fcPlcfFldTxbx lcbPlcfFldTxbx fcPlcfHdrtxbxTxt lcbPlcfHdrtxbxTxt fcPlcffldHdrTxdx lcbPlcffldHdrTxdx fcStwUser lcbStwUser fcSttbTtmbd lcbSttbTtmbd fcCookieData lcbCookieData fcPgdMotherOldOld lcbPgdMotherOldOld fcBkdMotherOldOld lcbBkdMotherOldOld fcPgdFtnOldOld lcbPgdFtnOldOld fcBkdFtnOldOld lcbBkdFtnOldOld fcPgdEdnOldOld lcbPgdEdnOldOld fcBkdEdnOldOld lcbBkdEdnOldOld fcSttbfIntlFld lcbSttbfIntlFld fcRouteSlip lcbRouteSlip fcSttbSavedBy lcbSttbSavedBy fcSttbFnm lcbSttbFnm fcPlfLst lcbPlfLst fcPlfLfo lcbPlfLfo fcPlcTxbxBkd lcbPlcTxbxBkd fcPlcfTxbxHdrBkd lcbPlcfTxbxHdrBkd fcDocUndoWord9 lcbDocUndoWord9 fcRgbUse lcbRgbUse fcUsp lcbUsp fcUskf lcbUskf fcPlcupcRgbUse lcbPlcupcRgbUse fcPlcupcUsp lcbPlcupcUsp fcSttbGlsyStyle lcbSttbGlsyStyle fcPlgosl lcbPlgosl fcPlcocx lcbPlcocx fcPlcfBteLvc lcbPlcfBteLvc dwLowDateTime dwHighDateTime fcPlcfLvcPe10 lcbPlcfLvcPe10 fcPlcfAsumy lcbPlcfAsumy fcPlcfGram lcbPlcfGram fcSttbListNames lcbSttbListNames fcSttbfUssr lcbSttbfUssr')
RULE_FIBFCLCB_PATTERN_97 = '=186L'

SIZE_OF_FIBFCLCB_2000 = 0x78      # 120Bytes
RULE_FIBFCLCB_NAME_2000 = ('fcPlcfTch lcbPlcfTch fcRmdThreading lcbRmdThreading fcMid lcbMid fcSttbRgtplc lcbSttbRgtplc fcMsoEnvelope lcbMsoEnvelope fcPlcfLad lcbPlcfLad fcRgDofr lcbRgDofr fcPlcosl lcbPlcosl fcPlcfCookieOld lcbPlcfCookieOld fcPgdMotherOld lcbPgdMotherOld fcBkdMotherOld lcbBkdMotherOld fcPgdFtnOld lcbPgdFtnOld fcBkdFtnOld lcbBkdFtnOld fcPgdEdnOld lcbPgdEdnOld fcBkdEdnOld lcbBkdEdnOld')
RULE_FIBFCLCB_PATTERN_2000 = '=30L'

SIZE_OF_FIBFCLCB_2002 = 0xE0      # 224Bytes
RULE_FIBFCLCB_NAME_2002 = ('fcUnused1 lcbUnused1 fcPlcfPgp lcbPlcfPgp fcPlcfuim lcbPlcfuim fcPlfguidUim lcbPlfguidUim fcAtrdExtra lcbAtrdExtra fcPlrsid lcbPlrsid fcSttbfBkmkFactoid lcbSttbfBkmkFactoid fcPlcfBkfFactoid lcbPlcfBkfFactoid fcPlcfcookie lcbPlcfcookie fcPlcfBklFactoid lcbPlcfBklFactoid fcFactoidData lcbFactoidData fcDocUndo lcbDocUndo fcSttbfBkmkFcc lcbSttbfBkmkFcc fcPlcfBkfFcc lcbPlcfBkfFcc fcPlcfBklFcc lcbPlcfBklFcc fcSttbfbkmkBPRepairs lcbSttbfbkmkBPRepairs fcPlcfbkfBPRepairs lcbPlcfbkfBPRepairs fcPlcfbklBPRepairs lcbPlcfbklBPRepairs fcPmsNew lcbPmsNew fcODSO lcbODSO fcPlcfpmiOldXP lcbPlcfpmiOldXP fcPlcfpmiNewXP lcbPlcfpmiNewXP fcPlcfpmiMixedXP lcbPlcfpmiMixedXP fcUnused2 lcbUnused2 fcPlcffactoid lcbPlcffactoid fcPlcflvcOldXP lcbPlcflvcOldXP fcPlcflvcNewXP lcbPlcflvcNewXP fcPlcflvcMixedXP lcbPlcflvcMixedXP')
RULE_FIBFCLCB_PATTERN_2002 = '=56L'

SIZE_OF_FIBFCLCB_2003 = 0xE0      # 224Bytes
RULE_FIBFCLCB_NAME_2003 = ('fcHplxsdr lcbHplxsdr fcSttbfBkmkSdt lcbSttbfBkmkSdt fcPlcfBkfSdt lcbPlcfBkfSdt fcPlcfBklSdt lcbPlcfBklSdt fcCustomXForm lcbCustomXForm fcSttbfBkmkProt lcbSttbfBkmkProt fcPlcfBkfProt lcbPlcfBkfProt fcPlcfBklProt lcbPlcfBklProt fcSttbProtUser lcbSttbProtUser fcUnused lcbUnused fcPlcfpmiOld lcbPlcfpmiOld fcPlcfpmiOldInline lcbPlcfpmiOldInline fcPlcfpmiNew lcbPlcfpmiNew fcPlcfpmiNewInline lcbPlcfpmiNewInline fcPlcflvcOld lcbPlcflvcOld fcPlcflvcOldInline lcbPlcflvcOldInline fcPlcflvcNew lcbPlcflvcNew fcPlcflcxNewInline lcbPlcflcxNewInline fcPgdMother lcbPgdMother fcBkdMother lcbBkdMother fcAfdMother lcbAfdMother fcPgdFtn lcbPgdFtn fcBkdFtn lcbBkdFtn fcAfdFtn lcbAfdFtn fcPgdEdn lcbPgdEdn fcBkdEdn lcbBkdEdn fcAfdEdn lcbAfdEdn fcAfd lcbAfd')
RULE_FIBFCLCB_PATTERN_2003 = '=56L'

SIZE_OF_FIBFCLCB_2007 = 0x98      # 152Bytes
RULE_FIBFCLCB_NAME_2007 = ('fcPlcfmthd lcbPlcfmthd fcSttbfBkmkMoveFrom lcbSttbfBkmkMoveFrom fcPlcfBkfMoveFrom lcbPlcfBkfMoveFrom fcPlcfBklMoveFrom lcbPlcfBklMoveFrom fcSttbfBkmkMoveTo lcbSttbfBkmkMoveTo fcPlcfBkfMoveTo lcbPlcfBkfMoveTo fcPlcfBklMoveTo lcbPlcfBklMoveTo fcUnused1 lcbUnused1 fcUnused2 lcbUnused2 fcUnused3 lcbUnused3 fcSttbfBkmkArto lcbSttbfBkmkArto fcPlcfBkfArto lcbPlcfBkfArto fcPlcfBklArto lcbPlcfBklArto fcArtoData lcbArtoData fcUnused4 lcbUnused4 fcUnused5 lcbUnused5 fcUnused6 lcbUnused6 fcOssTheme lcbOssTheme fcColorSchemeMapping lcbColorSchemeMapping')
RULE_FIBFCLCB_PATTERN_2007 = '=38L'

class CHWP():
    def fnFindHWPHeader(self, s_buffer, dl_OLEDirectory, t_SAT, t_SSAT):
        t_HWPHeader = ()
        
        try :
            
            StructOLE = CStructOLE()
            StructHWP = CStructHWP()
            s_RootEntry = StructOLE.fnStructOLESectorEx(s_buffer, None, dl_OLEDirectory, t_SAT, t_SSAT, "RootEntry")
            if s_RootEntry == "" :
                print "[-] Error - StructHWPHeader( RootEntry )"
                return t_HWPHeader
            
            s_HWPHeader = StructOLE.fnStructOLESectorEx(s_buffer, s_RootEntry, dl_OLEDirectory, t_SAT, t_SSAT, "FileHeader")
            if s_HWPHeader == "" :
                print "[-] Error - StructHWPHeader( FileHeader )"
                return t_HWPHeader
            
            t_HWPHeader = StructHWP.fnStructHWPHeader(s_HWPHeader)
            if t_HWPHeader == () :
                print "[-] Error - fnStructHWPHeader()"
                return t_HWPHeader
            
        except :
            print format_exc()
            
        return t_HWPHeader

class CStructHWP():
    def fnStructHWPHeader(self, s_HWPHeader):
        t_HWPHeader = ()
        
        try :
            
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

class CExceptHWP(CStructHWP):
    pass

RULE_HWPHEADER_NAME = ('Signature Version Property Reserved')
RULE_HWPHEADER_PATTERN = '=32s2l216s'

class CPrintOLE():
    def fnPrintOLEHeader(self):
        pass
    def fnPrintOLEDirectory(self, dl_OLEDirectory):
        
        try :
            UtilOLE = CUtilOLE()
            
            print "\t" + "=" * 130
            print "\t" + "%-33s%-11s%-11s%-11s%-11s%-11s%-11s%-11s%-11s%-10s" % ("DirectoryName", "szName", "Type", "Color", "Left", "Right", "Root", "Flags", "SecID", "szStream")
            print "\t" + "-" * 130
            n_Cnt = 0
            while n_Cnt < dl_OLEDirectory.__len__() :
                if UtilOLE.fnExtractAlphaNumber(dl_OLEDirectory[n_Cnt][0]) == "" :
                    n_Cnt +=1
                    continue
                
                for n_Index in range( dl_OLEDirectory[n_Cnt].__len__() ) :
                    if n_Index == 0 :
                        print "\t" + "%-30s" % UtilOLE.fnExtractAlphaNumber(dl_OLEDirectory[n_Cnt][0]),
                    elif n_Index == 7 or n_Index == 9 or n_Index == 10 or n_Index == 13:
                        continue
                    else :
                        print "0x%08X" % dl_OLEDirectory[n_Cnt][n_Index],
                
                print ""
                n_Cnt += 1
            
            print "\t" + "=" * 130 + "\n"
            
        except :
            print format_exc()
            return False
        
        return True
        
class CPrintOffice():
    def fnPrintOfficeStream(self, dl_StreamOffset, dl_StreamName, dl_StreamSize):
        
        try :
            l_Version = ["97", "2000", "2002", "2003", "2007"]
            
            print "\n\t\t\t" + "=" * 50
            print "\t\t\t%4s%19s\t%10s\t%7s" % ("Index", "Member", "Offset", "Size")
            print "\t\t\t" + "-" * 50
            for n_Ver in range( dl_StreamOffset.__len__() ) :
                print "\t\t\t[ Office%4s ]" % l_Version[ n_Ver ]
                for n_Cnt in range( dl_StreamOffset[n_Ver].__len__() ) :
                    print "\t\t\t%02X%22s\t0x%08X\t0x%08X" % (n_Cnt, dl_StreamName[n_Ver][n_Cnt], dl_StreamOffset[n_Ver][n_Cnt], dl_StreamSize[n_Ver][n_Cnt])
            print "\t\t\t" + "=" * 50 + "\n"
            
        except :
            print format_exc()
            return False
        
        return True

class CPrintHWP():
    pass






