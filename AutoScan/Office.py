
from struct import unpack
from traceback import format_exc
from collections import namedtuple

from Definition import *
from Common import CBuffer, CFile
from OLE_Stream import CMappedStream


class COffice():
    
    #    s_pBuf                        [IN]                    File Full Buffer
    #    t_SATList                     [IN]                    SSAT List from SAT
    #    l_OLEDirectory                [IN]                    OLE Directory List ( Double-Array )
    #    BOOL Type                     [OUT]                   True / False
    @classmethod
    def fnScan(cls, s_fname, s_pBuf, t_SATList, t_SSATList, dl_OLEDirectory):
        
        print "\t\t[+] COffice.Scan( )"
        
        try :

# Step 1. Extract WordDocument 
            StructOffice = CStructOffice()
            
            s_WordDoc = ""
            s_WordDoc = StructOffice.fnStructOfficeSector(s_pBuf, s_WordDoc, dl_OLEDirectory, t_SATList, t_SSATList, "WordDocument")
            if s_WordDoc == None :
                print "\t\t\t[-] Failure - StructOfficeSector( WordDocument )"
                return False
            
            CFile.fnWriteFile("c:\\test1\\%s_WordDocument.dump" % s_fname, s_WordDoc)
            
#            n_Cnt = 0
#            while n_Cnt < dl_OLEDirectory.__len__() :
#                print CBuffer.fnExtractAlphaNumber(dl_OLEDirectory[n_Cnt][0])
#                n_Cnt += 1

# Step 2. Mapped WordDocument
            t_FibBase, l_FibRgFcLcbBlob = StructOffice.fnStructOfficeWordDocument(s_WordDoc)
            if t_FibBase == None or l_FibRgFcLcbBlob == None :
                print "\t\t\t[-] Failure - MappedOfficeWordDocument()"
                return False

#            print "\t\t\t[*] File Version : Office%s" % g_Version[ l_FibRgFcLcbBlob.__len__() - 1 ]


# Step 3. Check Encryption & Obfuscation
            s_Algorithm = StructOffice.fnStructOfficeFlag(t_FibBase.Flag)
            if s_Algorithm == None or s_Algorithm == "" :
                print "\t\t\t[-] Failure - StructOfficeFlag( 0x%08X )" % t_FibBase.Flag
                return False
            
            if  s_Algorithm == "RC4" or s_Algorithm == "XOR" :
                print "\t\t\t[-] Failure - Do not Support! ( %s )" % s_Algorithm
                return False
            

# Step 4. Mapped Table Stream
            b_Bit = CBuffer.fnBitParse( t_FibBase.Flag, 9, 1 )
            if b_Bit == None :
                print "\t\t\t[-] Error - BitParse( Table Name, BitCount : 9 )"
                return False
            elif b_Bit == 1 :
                s_StreamName = "1Table"
            else :
                s_StreamName = "0Table"
                
            s_TableStream = StructOffice.fnStructOfficeSector(s_pBuf, None, dl_OLEDirectory, t_SATList, t_SSATList, s_StreamName)
            if s_TableStream == None :
                print "\t\t\t[-] Failure - StructOfficeSector( %s )" % s_StreamName
                return False
                      

# Step 5. Mapped Streams ( WordDocument, Table Stream, Data Stream, FibRgFcLcbBlob List )
            # it is 'Possible' that s_DataStream was None
            s_DataStream = StructOffice.fnStructOfficeSector(s_pBuf, None, dl_OLEDirectory, t_SATList, t_SSATList, "Data")
            
#            CFile.fnWriteFile("c:\\test1\\%s_Data.dump" % s_fname, s_DataStream)

            StructStream = CStructStream()
            if not StructStream.fnStructOfficeStream(s_fname, s_WordDoc, s_TableStream, s_DataStream, l_FibRgFcLcbBlob) :
                print "\t\t\t[-] Failure - StructOfficeStream()"
                return False
                    
        except :
            print format_exc()
            return False

        return True
        
        
        
class CStructOffice():
    
    #    s_pBuf                             [IN]                        File Full Buffer
    #    s_WordDoc                          [IN]                        WordDocument Buffer
    #    dl_OLEDirectory                    [IN]                        Double-List OLEDirectory
    #    t_SATList                          [IN]                        SAT List
    #    t_SSATList                         [IN]                        SSAT List
    #    s_Entry                            [IN]                        Targeted Entry Name
    #    s_Sector                           [OUT]                       Sector Data By EntryName    
    def fnStructOfficeSector(self, s_pBuf, s_pBuff, dl_OLEDirectory, t_SATList, t_SSATList, s_Entry):
        
        from OLE import CStructOLE
        
        s_Sector = ""
        
        try :
            
            StructOLE = CStructOLE()
            
            s_EntryName, n_Index = StructOLE.fnFindEntryName(dl_OLEDirectory, s_Entry)
            if s_EntryName == None :
                print "\t\t\t[-] Failure - FindEntryName() for %s in Office" % s_Entry
                return None
            
            n_EntryType  = dl_OLEDirectory[n_Index][2]           # Entry Sector Type
            n_EntrySecID = dl_OLEDirectory[n_Index][11]          # Entry Sector Start SecID
            n_EntrySize  = dl_OLEDirectory[n_Index][12]          # Entry Sector Size
            
            if n_EntryType == 5 or n_EntrySize >= 0x1000 :
                s_Buffer = s_pBuf
                t_Table = t_SATList
                n_Size = SIZE_OF_SECTOR
            elif n_EntryType != 5 and n_EntrySize < 0x1000 :
                if s_pBuff == None :
                    print "\t\t\t[-] Error - StructOfficeSector()'s Parameter ( Entry Name : %s, Type : 0x%08X )" % (s_Entry, n_EntryType)
                    return None
                
                s_Buffer = s_pBuff
                t_Table = t_SSATList
                n_Size = SIZE_OF_SHORT_SECTOR
            else :
                print "\t\t\t[-] Error - Do Not Seperated referred Type"
                return None
            
            s_Sector = StructOLE.fnStructOLESector(s_Buffer, t_Table, n_EntrySecID, n_Size)
            if s_Sector == None :
                print "\t\t\t[-] Failure - StructOLESector() for %s in Office" % s_Entry
                return None
            
        except :
            print format_exc()
            return None
        
        return s_Sector[0:n_EntrySize]

    
    #    s_WordDoc                                [IN]                    WordDocument Sector's Buffer
    #    l_FibRgFcLcbBlob                         [OUT]                   Mapped FibRgFcLcbBlob List
    def fnStructOfficeWordDocument(self, s_WordDoc):
        
        try :
            MappedOffice = CMappedOffice()

            n_Position = 0

# Step 1. Mapped FibBase ( Position : 0x0000,  Size : 0x0020 )
            t_FibBase = MappedOffice.fnMappedFibBase(s_WordDoc[n_Position:n_Position + SIZE_OF_FIBBASE])
            if t_FibBase == None :
                print "\t\t\t[-] Failure - MappedFibBase()"
                return None, None
            
            # Check Signature
            if t_FibBase.wIdent != 0xA5EC :
                print "\t\t\t[-] Error - FibBase's Signature( 0x%08X )" % t_FibBase.wIdent
                return None, None
                    
            n_Position += SIZE_OF_FIBBASE
            
            
# Step 2. Mapped FibRgW97 
            # Get Size FibRg97 ( Position : 0x0020, Size : 0x0002 )
            n_CSW = CBuffer.fnReadWord(s_WordDoc, n_Position)
            if n_CSW != 0x000E :
                print "\t\t\t[-] Error ( Position : 0x%08X, CSW : 0x%04X )" % (n_Position, n_CSW)
                return None, None
            
            n_Position += 2
            
            # Mapped FibRgW97 ( Position : 0x0022, Size : 0x001C )
            n_szCSW = n_CSW * 2            
            t_FibW97 = MappedOffice.fnMappedFibRgW97(s_WordDoc[n_Position:n_Position + n_szCSW])
            if t_FibW97 == None :
                print "\t\t\t[-] Failure - MappedFibRgW97( Position : 0x%08X, CSW : 0x%04X )" % (n_Position, n_CSW)
                return None, None
            
            n_Position += n_szCSW
            
            
# Step 3. Mapped FibRgLw97
            # Get Size FibRglW97 ( Position : 0x003E, Size : 0x0002 )
            n_CSLW = CBuffer.fnReadWord(s_WordDoc, n_Position)    
            if n_CSLW != 0x0016 :
                print "\t\t\t[-] Error ( Position : 0x%08X, CSLW : 0x%04X )" % (n_Position, n_CSLW)
                return None, None
            
            n_Position += 2
            
            # Mapped FibRglW97 ( Position : 0x0040, Size : 0x0058 )
            n_szCSLW = n_CSLW * 4
            t_FibLw97 = MappedOffice.fnMappedFibRgLw97(s_WordDoc[n_Position:n_Position + n_szCSLW])
            if t_FibLw97 == None :
                print "\t\t\t[-] Failure - MappedFibRgLw97( Position : 0x%08X, CSLW : 0x%04X )" % (n_Position, n_CSLW)
                return None, None
            
            n_Position += n_szCSLW
            
            
# Step 4. Mapped cbRgFcLcb
            # Get Size cbRgFcLcb ( Position : 0x0098, Size : 0x0002 )
            n_cbRgFcLcb = CBuffer.fnReadWord(s_WordDoc, n_Position)
            if not n_cbRgFcLcb in g_cbRgFcLcb :            
                print "\t\t\t[-] Error ( Position : 0x%08X, cbRgFcLcb : 0x%04X )" % (n_Position, n_cbRgFcLcb)
                return None, None
                
            n_Position += 2
            
            # Delay Mapped for Pre-Reading nFib Value
            n_Position_FibRgFcLcbBlob = n_Position
            
            n_Position += n_cbRgFcLcb * 8
            

# Step 5. Fixed nFib
            # Get cswNew & Checking 
            n_cswNew = CBuffer.fnReadWord(s_WordDoc, n_Position)
            if not n_cswNew in g_cswNew :
                print "\t\t\t[-] Error ( Position : 0x%08X, cswNew : 0x%04X )" % (n_Position, n_cswNew)
                return None, None
            
            n_Position += 2
            
            if n_cswNew == 0 :
                n_Fib = t_FibBase.nFib
            else :
                n_Fib = CBuffer.fnReadWord(s_WordDoc, n_Position)
            
            if not n_Fib in g_Fib :
                print "\t\t\t[-] Error ( cswNew : 0x%04X, nFib : 0x%04X )" % (n_cswNew, n_Fib)
                return None, None
            
            n_Position_CswNewData = n_Position                


# Step 6. Mapped FibRgFcLcbBlob
            n_Position = n_Position_FibRgFcLcbBlob
            l_FibRgFcLcbBlob = MappedOffice.fnMappedFibRgFcLcbBlob(s_WordDoc[n_Position:], n_Fib)
            if l_FibRgFcLcbBlob == None :
                print "\t\t\t[-] Failure - MappedFibRgFcLcbBlob( Position : 0x%08X, nFib : 0x%04X )" % (n_Position, n_Fib)
                return None, None

        except :
            print format_exc()
            return None, None
        
        return t_FibBase, l_FibRgFcLcbBlob       
    
    
    #    n_Flag                                [IN]                        t_FibBase.Flag Data
    #    s_Algorithm                           [OUT]                       Algorithm ( XOR, RC4, None )
    def fnStructOfficeFlag(self, n_Flag):
        
        try :
            
            s_Algorithm = ""
            
            # Get Encrypted Bit
            b_Bit_Encrypted = CBuffer.fnBitParse(n_Flag, 8, 1)
            if b_Bit_Encrypted == None :
                print "\t\t\t[-] Error - BitParse( Encrypted, BitCount : 8 )"
                return None
            elif b_Bit_Encrypted == 1 :
                n_Encrypted = 1
            else :
                n_Encrypted = 0

            # Get Obfuscated Bit
            b_Bit_Obfuscated = CBuffer.fnBitParse(n_Flag, 15, 1)
            if b_Bit_Obfuscated == None :
                print "\t\t\t[-] Error - BitParse( Obfuscated, BitCount : 15 )"
                return None
            elif b_Bit_Obfuscated == 1 :
                n_Obfuscated = 1
            else :
                n_Obfuscated = 0
                
            # Check Algorithm
            if n_Encrypted == 0 and n_Obfuscated == 0 :
                s_Algorithm = "None"
            
            if n_Encrypted == 1 and n_Obfuscated == 0 :
                s_Algorithm = "RC4"
            
            if n_Encrypted == 1 and n_Obfuscated == 1 :
                s_Algorithm = "XOR"
            
        except :
            print format_exc()
            return None
        
        return s_Algorithm
    
    
    
    
    


class CMappedOffice():

    #    s_Buffer                                 [IN]                    FibBase Sector's Buffer
    #    t_FibBase                                [OUT]                   Mapped FibBase
    def fnMappedFibBase(self, s_Buffer):
        
        try :
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibBase.dump", s_Buffer)
            
            t_FibBase_Name = namedtuple('FibBase', RULE_FIBBASE_NAME)
            t_FibBase = t_FibBase_Name._make( unpack(RULE_FIBBASE_PATTERN, s_Buffer) )
            if t_FibBase == () :
                return None           
            
        except : 
            print format_exc()
            return None
        
        return t_FibBase


    #    s_Buffer                                 [IN]                    FibRgW97 Sector's Buffer
    #    t_FibW97                                 [OUT]                   Mapped Struct FibRgW97
    def fnMappedFibRgW97(self, s_Buffer):
        
        try :
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibRgW97.dump", s_Buffer)
        
            t_FibW97_Name = namedtuple('FibW97', RULE_FIBW97_NAME)
            t_FibW97 = t_FibW97_Name._make( unpack(RULE_FIBW97_PATTERN, s_Buffer) )
            if t_FibW97 == () :
                return None            
            
        except :
            print format_exc()
            return None
        
        return t_FibW97


    #    s_Buffer                                 [IN]                    FibRgLw97 Sector's Buffer
    #    t_FibLw97                                [OUT]                   Mapped Struct FibRgLw97
    def fnMappedFibRgLw97(self, s_Buffer):
        
        try :
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibRgLw97.dump", s_Buffer)
            
            t_FibLw97_Name = namedtuple('FibLw97', RULE_FIBLW97_NAME)
            t_FibLw97 = t_FibLw97_Name._make( unpack(RULE_FIBLW97_PATTERN, s_Buffer) )
            if t_FibLw97 == () :
                return None
            
        except :
            print format_exc()
            return None
        
        return t_FibLw97


    #    s_Buffer                         [IN]                        FibRgFcLcbBlob Buffer
    #    n_Fib                            [IN]                        nFib
    #    l_FibRgFcLcb                     [OUT]                       Mapped FibRgFcLcb List
    def fnMappedFibRgFcLcbBlob(self, s_Buffer, n_Fib):
        
        l_FibRgFcLcb = []
        
        try :
            n_Position = 0
            
            # FibRgFcLcb97
            t_FibRgFcLcb97 = self.fnMappedFibRgFcLcb97(s_Buffer[n_Position:n_Position + SIZE_OF_FIBFCLCB_97])
            if t_FibRgFcLcb97 == None :
                print "\t\t\t[-] Failure - FibRgFcLcb97( Position : 0x%08X, Size : 0x%08X )" % (n_Position, SIZE_OF_FIBFCLCB_97)
                return None
            
            l_FibRgFcLcb.append( t_FibRgFcLcb97 )
            if n_Fib == 0x00C1 :
                return l_FibRgFcLcb
            n_Position += SIZE_OF_FIBFCLCB_97
            
            
            # FibRgFcLcb2000
            t_FibRgFcLcb2000 = self.fnMappedFibRgFcLcb2000(s_Buffer[n_Position:n_Position + SIZE_OF_FIBFCLCB_2000])
            if t_FibRgFcLcb2000 == None :
                print "\t\t\t[-] Failure - FibRgFcLcb2000( Position : 0x%08X, Size : 0x%08X )" % (n_Position, SIZE_OF_FIBFCLCB_2000)
                return None
            
            l_FibRgFcLcb.append( t_FibRgFcLcb2000 )
            if n_Fib == 0x00D9 :
                return l_FibRgFcLcb
            n_Position += SIZE_OF_FIBFCLCB_2000
            
            
            # FibRgFcLcb2002
            t_FibRgFcLcb2002 = self.fnMappedFibRgFcLcb2002(s_Buffer[n_Position:n_Position + SIZE_OF_FIBFCLCB_2002])
            if t_FibRgFcLcb2002 == None :
                print "\t\t\t[-] Failure - FibRgFcLcb2002( Position : 0x%08X, Size : 0x%08X )" % (n_Position, SIZE_OF_FIBFCLCB_2002)
                return None

            l_FibRgFcLcb.append( t_FibRgFcLcb2002 )
            if n_Fib == 0x0101 :
                return l_FibRgFcLcb
            n_Position += SIZE_OF_FIBFCLCB_2002
            
            
            # FibRgFcLcb2003
            t_FibRgFcLcb2003 = self.fnMappedFibRgFcLcb2003(s_Buffer[n_Position:n_Position + SIZE_OF_FIBFCLCB_2003])
            if t_FibRgFcLcb2003 == None :
                print "\t\t\t[-] Failure - FibRgFcLcb2003( Position : 0x%08X, Size : 0x%08X )" % (n_Position, SIZE_OF_FIBFCLCB_2003)
                return None
            
            l_FibRgFcLcb.append( t_FibRgFcLcb2003 )
            if n_Fib == 0x010C :
                return l_FibRgFcLcb            
            n_Position += SIZE_OF_FIBFCLCB_2003
            
            
            # FibRgFcLcb2007
            t_FibRgFcLcb2007 = self.fnMappedFibRgFcLcb2007(s_Buffer[n_Position:n_Position + SIZE_OF_FIBFCLCB_2007])
            if t_FibRgFcLcb2007 == None :
                print "\t\t\t[-] Failure - FibRgFcLcb2007( Position : 0x%08X, Size : 0x%08X )" % (n_Position, SIZE_OF_FIBFCLCB_2007)
                return None
            
            l_FibRgFcLcb.append( t_FibRgFcLcb2007 )
            
            if n_Fib != 0x0112 :
                print "\t\t\t[-] Error ( nFib : 0x%04X )" % n_Fib
                return None            
            
        except :
            print format_exc()
            return None
        
        return l_FibRgFcLcb
        

    #    s_Buffer                        [IN]                        FibRgFcLcb97 Buffer
    #    t_FibRgFcLcb97                  [OUT]                       Mapped FibRgFcLcb97
    def fnMappedFibRgFcLcb97(self, s_Buffer):
        
        try : 
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibRgFcLcb97.dump", s_Buffer)
            
            t_FibRgFcLcb97_Name = namedtuple('FibRgFcLcb97', RULE_FIBFCLCB_NAME_97)
            t_FibRgFcLcb97 = t_FibRgFcLcb97_Name._make( unpack(RULE_FIBFCLCB_PATTERN_97, s_Buffer))
            if t_FibRgFcLcb97 == () :
                return None
        
        except :
            print format_exc()
            return None
        
        return t_FibRgFcLcb97


    #    s_Buffer                        [IN]                        FibRgFcLcb97 Buffer
    #    t_FibRgFcLcb2000                [OUT]                       Mapped FibRgFcLcb2000
    def fnMappedFibRgFcLcb2000(self, s_Buffer):
        
        try : 
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibRgFcLcb2000.dump", s_Buffer)
            
            t_FibRgFcLcb2000_Name = namedtuple('FibRgFcLcb2000', RULE_FIBFCLCB_NAME_2000)
            t_FibRgFcLcb2000 = t_FibRgFcLcb2000_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2000, s_Buffer))
            if t_FibRgFcLcb2000 == () :
                return None
        
        except :
            print format_exc()
            return None
        
        return t_FibRgFcLcb2000
    
    
    #    s_Buffer                        [IN]                        FibRgFcLcb97 Buffer
    #    t_FibRgFcLcb2002                [OUT]                       Mapped FibRgFcLcb2002    
    def fnMappedFibRgFcLcb2002(self, s_Buffer):
        
        try : 
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibRgFcLcb2002.dump", s_Buffer)
            
            t_FibRgFcLcb2002_Name = namedtuple('FibRgFcLcb2002', RULE_FIBFCLCB_NAME_2002)
            t_FibRgFcLcb2002 = t_FibRgFcLcb2002_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2002, s_Buffer))
            if t_FibRgFcLcb2002 == () :
                return None
        
        except :
            print format_exc()
            return None
        
        return t_FibRgFcLcb2002


    #    s_Buffer                        [IN]                        FibRgFcLcb97 Buffer
    #    t_FibRgFcLcb2003                [OUT]                       Mapped FibRgFcLcb2003
    def fnMappedFibRgFcLcb2003(self, s_Buffer):
        
        try : 
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibRgFcLcb2003.dump", s_Buffer)
            
            t_FibRgFcLcb2003_Name = namedtuple('FibRgFcLcb2003', RULE_FIBFCLCB_NAME_2003)
            t_FibRgFcLcb2003 = t_FibRgFcLcb2003_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2003, s_Buffer))
            if t_FibRgFcLcb2003 == () :
                return None
        
        except :
            print format_exc()
            return None
        
        return t_FibRgFcLcb2003


    #    s_Buffer                        [IN]                        FibRgFcLcb97 Buffer
    #    t_FibRgFcLcb2007                [OUT]                       Mapped FibRgFcLcb2007
    def fnMappedFibRgFcLcb2007(self, s_Buffer):
        
        try : 
            
#            CFile.fnWriteFile("c:\\test1\\1.doc_FibRgFcLcb2007.dump", s_Buffer)
            
            t_FibRgFcLcb2007_Name = namedtuple('FibRgFcLcb2007', RULE_FIBFCLCB_NAME_2007)
            t_FibRgFcLcb2007 = t_FibRgFcLcb2007_Name._make( unpack(RULE_FIBFCLCB_PATTERN_2007, s_Buffer))
            if t_FibRgFcLcb2007 == () :
                return None
        
        except :
            print format_exc()
            return None
        
        return t_FibRgFcLcb2007



class CStructStream():

    #    s_fname                                  [IN]                        File Full Name
    #    s_WordDoc                                [IN]                        WordDocument Stream Buffer
    #    s_Table                                  [IN]                        Table Stream ( 0Table or 1Table )
    #    s_Data                                   [IN]                        Data Stream
    #    l_FibRgFcLcbBlob                         [IN]                        FibRgFcLcb List in FIB
    #    BOOL Type                                [OUT]                       True / False
    def fnStructOfficeStream(self, s_fname, s_WordDoc, s_Table, s_Data, l_FibRgFcLcbBlob):
        
        try :

# Step 1. Get Structure List
            # Get Enable Structure Data by Structure's Member Size in l_FibRgFcLcbBlob
            dl_StreamOffset, dl_StreamSize, dl_StreamName = self.fnStructStreamData(l_FibRgFcLcbBlob, " ")
            if dl_StreamOffset == None or dl_StreamOffset == [] :
                print "\t\t\t[-] Failure - StructStreamData( Offset )"
                return False
            if dl_StreamSize == None or dl_StreamSize == [] :
                print "\t\t\t[-] Failure - StructStreamData( Size )"
                return False
            if dl_StreamName == None or dl_StreamName == [] :
                print "\t\t\t[-] Failure - StructStreamData( Name )"
                return False
        
        
#            print "\t\t\t" + "=" * 50
#            print "\t\t\t%4s%19s\t%10s\t%7s" % ("Index", "Member", "Offset", "Size")
#            print "\t\t\t" + "-" * 50
#            for n_Ver in range( dl_StreamOffset.__len__() ) :
#                print "\t\t\t[ Office%4s ]" % g_Version[ n_Ver ]
#                for n_Cnt in range( dl_StreamOffset[n_Ver].__len__() ) :
#                    print "\t\t\t%02X%22s\t0x%08X\t0x%08X" % (n_Cnt, dl_StreamName[n_Ver][n_Cnt], dl_StreamOffset[n_Ver][n_Cnt], dl_StreamSize[n_Ver][n_Cnt])
#            print "\t\t\t" + "=" * 50
        

# Step 2. Parse Structure List
            if not self.fnStructParseStream(s_fname, s_WordDoc, s_Table, dl_StreamName, dl_StreamOffset, dl_StreamSize) :
                print "\t\t\t[-] Failure - StructParseStream()"
                return False
        
        except :
            print format_exc()
            return False
        
        return True


    #    l_FibRgFcLcbBlob                        [IN]                        FibRgFcLcbBlob Lists
    #    dl_StreamOffset                         [OUT]                       Stream Offset Lists for ParseTarget
    #    dl_StreamSize                           [OUT]                       Stream Size Lists for ParseTarget
    #    dl_StreamName                           [OUT]                       Stream Name Lists for ParseTarget
    def fnStructStreamData(self, l_FibRgFcLcbBlob, s_Seperator):
        
        dl_StreamOffset = []
        dl_StreamSize = []
        dl_StreamName = []
        
        try :
            # Convert String2List By Seperator "Null Space"
            l_FibFcLcbName = []
            dl_FibFcLcbName = []
            for n_Ver in range( l_FibRgFcLcbBlob.__len__() ) :
                l_FibFcLcbName = CBuffer.fnStr2List( eval( "RULE_FIBFCLCB_NAME_" + g_Version[ n_Ver ] ), s_Seperator )
                dl_FibFcLcbName.append( l_FibFcLcbName )
               
            
            # Get Stream Data
            for n_Ver in range( l_FibRgFcLcbBlob.__len__() ) :            # FibRgFcLcb97 ~ FibRgFcLcb2007
                l_Fc = []
                l_Lcb = []
                l_Name = []
                
                
                for n_Cnt in range( l_FibRgFcLcbBlob[n_Ver].__len__() / 2 ) :
                    n_Fc = l_FibRgFcLcbBlob[n_Ver][2 * n_Cnt]
                    n_Lcb = l_FibRgFcLcbBlob[n_Ver][2 * n_Cnt + 1]
                    s_FcName = dl_FibFcLcbName[n_Ver][2 * n_Cnt]
                    s_LcbName = dl_FibFcLcbName[n_Ver][2 * n_Cnt + 1]
                    
                    if n_Lcb == 0 :
                        continue
                    
                    if s_FcName in g_ExceptList :
                        continue
                    
                    if s_FcName[:2].find("fc") != 0 or s_LcbName[:3].find("lcb") != 0 :
                        print "\t\t\t[-] Error (%s[Ret : %d], %s[Ret : %d])" % (s_FcName, s_FcName[:2].find("fc"), s_LcbName, s_LcbName[:3].find("lcb"))
                        return None, None, None
                    
                    l_Fc.append( n_Fc )
                    l_Lcb.append( n_Lcb )
                    l_Name.append( s_FcName[2:] )
                    
                    
                dl_StreamOffset.append( l_Fc )
                dl_StreamSize.append( l_Lcb )
                dl_StreamName.append( l_Name ) 
        
        except IndexError :
            print format_exc()
            print "n_Ver : 0x%04X, n_Cnt : 0x%04X" % (n_Ver, n_Cnt)
            return None, None, None
        
        except :
            print format_exc()
            return None, None, None
        
        return dl_StreamOffset, dl_StreamSize, dl_StreamName


    #    s_fname                                          [IN]                        File Full Name
    #    s_WordDoc                                        [IN]                        WordDocument Stream Buffer
    #    s_Table                                          [IN]                        Table Stream Buffer ( 0Table or 1Table )
    #    dl_Name                                          [IN]                        Member Name in Structure
    #    dl_Offset                                        [IN]                        Member Offset in Table Stream
    #    dl_Size                                          [IN]                        Member Size in Table Stream
    #    BOOL Type                                        [OUT]                       True / False
    def fnStructParseStream(self, s_fname, s_WordDoc, s_Table, dl_Name, dl_Offset, dl_Size):
        
        try :
            
            MappedStream = CMappedStream()
            
            for n_Ver in range( dl_Name.__len__() ) :
                for s_Name in dl_Name[ n_Ver ] :
                    if s_Name in g_StructList :
                        eval("MappedStream.fnMapped" + s_Name)( s_fname, s_WordDoc, s_Table, dl_Offset[ n_Ver ][ dl_Name[ n_Ver ].index( s_Name ) ], dl_Size[ n_Ver ][ dl_Name[ n_Ver ].index( s_Name )] )           
            
        except :
            print format_exc()
            return False

        return True
        


