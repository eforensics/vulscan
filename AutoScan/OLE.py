# -*- coding:utf-8 -*-


from struct import error, unpack
from traceback import format_exc
from collections import namedtuple
from zlib import decompress

from Definition import *
from Common import CBuffer, CFile
from HWP import CHwp
from Office import COffice


class COLE():
    
    #    s_pBuf                            [IN]                        File Full Buffer 
    #    BOOL                              [OUT]                       True / False
    @classmethod
    def fnCheck(cls, s_pBuf):
        
        try :
            # Case1. 0xe011cfd0, 0xe11ab1a1L
            if (s_pBuf[0:4] == "\xd0\xcf\x11\xe0" and s_pBuf[4:8] == "\xa1\xb1\x1a\xe1") or (s_pBuf[0:4] == "\xd0\xcd\x11\xe0") : 
                return "OLE"
            else :
                return None
            
        except error : 
            return None
        
        except :
            print format_exc()
            return None
        
    
    #    s_fname                            [IN]                         File Full Name
    #    s_pBuf                             [IN]                         File Full Buffer
    #    s_Format                           [IN]                         File Format
    #    BOOL Type                          [OUT]                        True / False
    @classmethod
    def fnScan(cls, s_fname, s_pBuf, s_Format):
        
        print "\t[+] Scan OLE - %s" % s_fname
        
        try :
            StructOLE = CStructOLE()
            
# Step 1. Parse OLE Header
            t_OLEHeader = StructOLE.fnStructOLEHeader(s_pBuf)
            if t_OLEHeader == None :
                print "\t\t[-] Failure - StructOLEHeader( %s )" % s_fname
                return False            
            
            
# Step 2. Extract OLE MSAT List
            t_MSATList = StructOLE.fnStructOLEMSAT(s_pBuf, t_OLEHeader.MSAT, t_OLEHeader.MSATSecID, t_OLEHeader.NumMSAT)
            if t_MSATList == None :
                print "\t\t[-] Failure - StructOLEMSAT( %s )" % s_fname
                return False
            
            # Check Count
            n_SATCnt = 0
            for n_SATIndex in t_MSATList :
                if n_SATIndex > 0 :
                    n_SATCnt += 1
            
            if n_SATCnt != t_OLEHeader.NumSAT :
                print "\t\t[-] Error - SAT Count( Index Count : 0x%08X, NumSAT : 0x%08X )" % (n_SATCnt, t_OLEHeader.NumSAT) 
                        
            
# Step 3. Extract OLE SAT List 
            t_SATList = StructOLE.fnStructOLESAT(s_pBuf, t_MSATList)
            if t_SATList == None :
                print "\t\t[-] Failure - StructOLESAT( %s )" % s_fname
                return False
            
            
# Step 4. Parse Directory Sector & Check File Format ( OLE-HWP5.x , OLE-Office )
            s_Directory = StructOLE.fnStructOLESector(s_pBuf, t_SATList, t_OLEHeader.DirSecID, SIZE_OF_SECTOR)
            if s_Directory == None :
                print "\t\t[-] Failure - StructOLESector( %s )" % s_fname
                return False
            
            dl_OLEDirectory = StructOLE.fnStructOLEDirectory(s_Directory)
            if dl_OLEDirectory == None :
                print "\t\t[-] Failure - StructOLEDirectory( %s )" % s_fname
                return False
            
            # Check Entry Name ( "HWP" or Not )
            s_Format, n_Index = StructOLE.fnFindEntryName(dl_OLEDirectory, "Hwp")
            if s_Format == None :
                s_Format = "Office"
            
            
# Step 5. Extract OLE SSAT List & Scan by Format
            t_SSATList = StructOLE.fnStructOLESSAT(s_pBuf, t_SATList, t_OLEHeader.SSATSecID, SIZE_OF_SECTOR )
            if t_SSATList == None :
                print "\t\t[-] Failure - StructOLESSAT( %s )" % s_fname
                return False
            
            eval( "C"+s_Format ).fnScan(s_fname, s_pBuf, t_SATList, t_SSATList, dl_OLEDirectory)
        
        except :
            print format_exc()
            return False
        
        return True



class CStructOLE():
    
    #    s_pBuf                            [IN]                File Full Buffer
    #    t_OLEHeader                       [OUT]               Mapped OLE Header
    def fnStructOLEHeader(self, s_pBuf):
        
        try :
            MappedOLE = CMappedOLE()
            t_OLEHeader = MappedOLE.fnMappedOLEHeader( s_pBuf[0:SIZE_OF_OLE_HEADER])
            
        except :
            print format_exc()
            
        return t_OLEHeader


    #    s_pBuf                            [IN]                        File Full Buffer
    #    s_MSAT                            [IN]                        MSAT in OLE File Header
    #    n_MSATSecID                       [IN]                        MSAT Sector ID in OLE File Header
    #    n_NumMSAT                         [IN]                        Number of MSAT in OLE File Header
    #    t_SAT                             [OUT]                       MSAT Tuple Data
    def fnStructOLEMSAT(self, s_pBuf, s_MSAT, n_MSATSecID, n_NumMSAT):
        
        try :
            # Check Parameters
            if n_NumMSAT < 0 :
                print "\t\t[-] Failure - StructOLEMSAT( )"
                return None            
            
            
            # MSAT List           
            t_MSAT = unpack("109l", s_MSAT)
            if n_NumMSAT != 0 :
                n_ExtraMSATCnt = n_NumMSAT
                n_ExtraMSecID = n_MSATSecID
                t_tmpMSAT = ()
                
                while True :
                    n_ExtraMSecPos = (n_ExtraMSecID + 1) * SIZE_OF_SECTOR
                    s_tmpMSAT = s_pBuf[n_ExtraMSecPos:n_ExtraMSecPos+SIZE_OF_SECTOR]
                    t_tmpMSAT = unpack("128l", s_tmpMSAT)
                    t_MSAT += t_tmpMSAT[:-1]
                    
                    n_ExtraMSecID = t_tmpMSAT[128-1]        # First Index is "0", Last Index is "MaxIndex -1"
                    
                    # Break Out Condition
                    n_ExtraMSATCnt -= 1
                    if n_ExtraMSATCnt == 0 :
                        break
                
        except :
            print format_exc()
            return None

        return t_MSAT


    #    s_pBuf                                   [IN]                File Full Buffer
    #    t_MSATList                               [IN]                SAT List from MSAT
    #    t_SAT                                    [OUT]               SSAT List from SAT
    def fnStructOLESAT(self, s_pBuf, t_MSATList):
    
        try :
            t_SAT = ()    
            
            # SAT List
            for n_SecID in t_MSATList :
                if n_SecID < 0 :
                    continue
                
                n_SecPos = (n_SecID + 1) * SIZE_OF_SECTOR
                s_tmpSAT = s_pBuf[n_SecPos:n_SecPos+SIZE_OF_SECTOR]
                t_tmpSAT = unpack("128l", s_tmpSAT)
                t_SAT += t_tmpSAT            
        
        except :
            print format_exc()
            return None
        
        return t_SAT
    
    
    #    s_pBuf                                [IN]                    File Full Buffer
    #    t_SATList                             [IN]                    SAT List
    #    SSATSecID                             [IN]                    SSAT Start SecID
    #    SIZE_OF_SECTOR                        [IN]                    Sector Size
    #    t_SSATList                            [OUT]                   SSAT List
    def fnStructOLESSAT(self, s_pBuf, SATList, SSATSecID, SIZE_OF_SECTOR):
        
        try :
            StructOLE = CStructOLE()            
            s_SSATList = StructOLE.fnStructOLESector(s_pBuf, SATList, SSATSecID, SIZE_OF_SECTOR)
            if s_SSATList == None :
                return None
            
            t_SSATList = unpack( str(s_SSATList.__len__() / 4)+"l", s_SSATList )
            
        except : 
            print format_exc()
            return None
        
        return t_SSATList


    #    s_pBuf                                [IN]                File Full Buffer
    #    t_Table                               [IN]                Table List
    #    n_SecID                               [IN]                Sector ID
    #    n_Size                                [IN]                Dump Size per Sector or SSector
    #    s_Sector                              [OUT]               Sector Data By n_SecID
    def fnStructOLESector(self, s_pBuf, t_Table, n_SecID, n_Size):

        s_Sector = ""
        
        try :
            
            if n_Size == SIZE_OF_SECTOR :
                AddSector = 1
            elif n_Size == SIZE_OF_SHORT_SECTOR :
                AddSector = 0
            
            while True :        
                n_SecPos = (n_SecID + AddSector) * n_Size
                s_Sector += s_pBuf[n_SecPos:n_SecPos+n_Size]
                n_SecID = t_Table[ n_SecID ]
                if n_SecID <= 0 :
                    break               
            
        except :
            print format_exc()
            return None
        
        return s_Sector

    
    #    s_Directory                            [IN]                Directory Buffer Strings
    #    t_Directory                            [OUT]               Directory Buffer Tuple
    def fnStructOLEDirectory(self, s_Directory):
        
        try :
            MappedOLE = CMappedOLE()
            
            if (s_Directory.__len__() % SIZE_OF_DIRECTORY) != 0 :
                print "\t\t\t[-] Error - OLEDirectory Size( %d )" % len(s_Directory)
                return None
            
            n_Start = 0
            l_Directory = []
            while True :
                s_SubDirectory = s_Directory[n_Start:n_Start + SIZE_OF_DIRECTORY]
                t_Directory = MappedOLE.fnMappedOLEDirectory(s_SubDirectory)
                l_Directory.append( list(t_Directory) )
                n_Start += SIZE_OF_DIRECTORY
                if n_Start == len(s_Directory) :
                    break
            
        except :
            print format_exc()
            return None

        return l_Directory


    #    dl_OLEDirectory                            [IN]                OLE Directory  ( Double-Array )
    #    s_CmpFormat                                [IN]                Find Format String
    #    s_Format                                   [OUT]               Found Format String or None
    #    n_Index                                    [OUT]               Found Format String by n_Index
    def fnFindEntryName(self, dl_OLEDirectory, s_CmpFormat):
        try :
            n_Index = 0
            s_Format = None
            
            while n_Index < dl_OLEDirectory.__len__() :
                s_Repaired = CBuffer.fnExtractAlphaNumber(dl_OLEDirectory[n_Index][0])
                if s_Repaired[:s_CmpFormat.__len__()].find( s_CmpFormat ) != -1 :
                    s_Format = s_CmpFormat
                    break
                n_Index += 1
                
        except :
            print format_exc()
            
        return s_Format, n_Index


class CMappedOLE():
    
    #    s_OLEHeader                        [IN]                File Buffer for OLEHeader ( Size by 0x200 )
    #    t_OLEHeader                        [OUT]               Mapped OLE Header
    def fnMappedOLEHeader(self, s_OLEHeader):
        
        try : 
            t_OLEHeader_Name = namedtuple('OLEHeader', RULE_OLEHEADER_NAME)
            t_OLEHeader = t_OLEHeader_Name._make( unpack(RULE_OLEHEADER_PATTERN, s_OLEHeader) )
            if t_OLEHeader == () : 
                return None
        
        except :
            print format_exc()
            return None
        
        return t_OLEHeader
            
    
    #    s_Directory                        [IN]                File Buffer for OLE Directory ( per 0x80 )  
    #    t_Directory                        [OUT]               Mapped OLE Directory
    def fnMappedOLEDirectory(self, s_Directory):
        
        try :
            
            t_OLEDirectory_Name = namedtuple('OLEDirectory', RULE_OLEDIRECTORY_NAME)
            t_OLEDirectory = t_OLEDirectory_Name._make( unpack(RULE_OLEDIRECTORY_PATTERN, s_Directory) )
            if t_OLEDirectory == () :
                return None
            
        except :
            print format_exc()
            return None
        
        return t_OLEDirectory













