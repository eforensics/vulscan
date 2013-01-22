# -*- coding:utf-8 -*-

import zlib
from zlib import decompress
from struct import unpack, error
from traceback import format_exc
from collections import namedtuple


from Definition import *
from Common import CFile, CBuffer
from PE import CPE



class CHwp23():
    
    #    s_pBuf                                [IN]                    File Full Buffer
    #    BOOL Type                             [OUT]                   True / False
    @classmethod
    def Check(cls, s_pBuf):
        
        try :
            s_Format = ""
            HWP23Sig = s_pBuf[0:0x20]
            if (HWP23Sig.find( "HWP Document File V2.00" ) != -1) or (HWP23Sig.find( "HWP Document File V3.00") != -1) :
                s_Format = "HWP23"
            else :
                s_Format = None
                
        except :
            print format_exc()
            return s_Format
        
        return s_Format
    

    @classmethod
    def Scan(cls, File):        
        pass
    


class CHwp():
    
    #    s_pBuf                        [IN]                    File Full Buffer
    #    t_SATList                     [IN]                    SSAT List from SAT
    #    dl_OLEDirectory               [IN]                    OLE Directory List ( Double-Array )
    #    BOOL Type                     [OUT]                   True / False
    @classmethod
    def fnScan(cls, s_fname, s_pBuf, t_SATList, t_SSATList, dl_OLEDirectory):
        
        print "\t\t[+] CHwp.Scan( )"
        
        try :
            
# Step 1. Extract RootEntry / FileHeader Sector
            StructHWP = CStructHWP()
            
            s_RootEntry = ""
            s_RootEntry = StructHWP.fnStructHWPSector(s_pBuf, None, dl_OLEDirectory, t_SATList, t_SSATList, "RootEntry")
            if s_RootEntry == None :
                print "\t\t\t[-] Failure - StructHWPSector( RootEntry )"
                return False
            
            s_FileHeader = StructHWP.fnStructHWPSector(s_pBuf, s_RootEntry, dl_OLEDirectory, t_SATList, t_SSATList, "FileHeader")
            if s_FileHeader == None :
                print "\t\t\t[-] Failure - StructHWPSector( FileHeader )"
                return False
            
            
# Step 2. Mapped "FileHeader" / Check Password
            MappedHWP = CMappedHWP() 
            
            t_FileHeader = MappedHWP.fnMappedHWPHeader(s_FileHeader)
            if t_FileHeader == None :
                print "\t\t\t[-] Failure - MappedHWPHeader()"
                return False
            
            if t_FileHeader.Property & 2 :
                print "\t\t\t[-] Failure - Password is required"
                print t_FileHeader
                return False
            
            
# Step 3. Extract HWP's Remind Sectors
            l_SSector = StructHWP.fnStructHWPSSector(s_RootEntry, dl_OLEDirectory, t_SSATList, t_FileHeader.Property & 1)
            if l_SSector == None :
                print "\t\t\t[-] Failure - StructHWPSSector() for Short-Sector in HWP"
                return False


# Step 4. Check Vulnerability
            ExploitHWP = CExploitHWP()
            if ExploitHWP.fnVulnerability(l_SSector) == None :
                print "\t\t\t[-] Failure - ExploitVulnerability()"
                return False
            
        except :
            print format_exc()
            return False
        
        return True
        
        
class CStructHWP():
    
    #    s_pBuf                             [IN]                        File Full Buffer
    #    s_RootEntry                        [IN]                        RootEntry Buffer
    #    dl_OLEDirectory                    [IN]                        Double-List OLEDirectory
    #    t_SATList                          [IN]                        SAT List
    #    t_SSATList                         [IN]                        SSAT List
    #    s_Entry                            [IN]                        Targeted Entry Name
    #    s_Sector                           [OUT]                       Sector Data By EntryName
    def fnStructHWPSector(self, s_pBuf, s_RootEntry, dl_OLEDirectory, t_SATList, t_SSATList, s_Entry ):
        
        from OLE import CStructOLE
        
        s_Sector = ""
        
        try :
            StructOLE = CStructOLE()
            
            s_EntryName, n_Index = StructOLE.fnFindEntryName(dl_OLEDirectory, s_Entry)
            if s_EntryName == None :
                print "\t\t\t[-] Failure - FindEntryName() for %s in HWP" % s_Entry
                return None
            
            n_EntryType  = dl_OLEDirectory[n_Index][2]           # Entry Sector Type
            n_EntrySecID = dl_OLEDirectory[n_Index][11]          # Entry Sector Start SecID
            n_EntrySize  = dl_OLEDirectory[n_Index][12]          # Entry Sector Size
            
            if n_EntryType == 5 or n_EntrySize >= 0x1000 :
                s_Buffer = s_pBuf
                t_Table = t_SATList
                n_Size = SIZE_OF_SECTOR
            elif n_EntryType != 5 and n_EntrySize < 0x1000 :
                if s_RootEntry == None :
                    print "\t\t\t[-] Error - StructHWPSector()'s Parameter ( Entry Name : %s, Type : 0x%08X )" % (s_Entry, n_EntryType)
                    return None
                
                s_Buffer = s_RootEntry
                t_Table = t_SSATList
                n_Size = SIZE_OF_SHORT_SECTOR
            else :
                print "\t\t\t[-] Error - Do Not Seperated referred Type"
                return None
            
            s_Sector = StructOLE.fnStructOLESector(s_Buffer, t_Table, n_EntrySecID, n_Size)
            if s_Sector == None :
                print "\t\t\t[-] Failure - StructOLESector() for %s in HWP" % s_Entry
                return None
        
        except :
            print format_exc()
            return None
        
        return s_Sector[0:n_EntrySize]
        
        

    #    s_RootEntry                         [IN]                   Short-Sector Based Buffer
    #    dl_OLEDirectory                     [IN]                   OLE Directory
    #    t_SSATList                          [IN]                   SSAT List
    #    B_CompressFlag                      [IN]                   BOOL Type's Compression Flag
    #    d_SSector                           [OUT]                  Extracted Short-Sectors
    def fnStructHWPSSector(self, s_RootEntry, dl_OLEDirectory, t_SSATList, B_CompressFlag) :
        
        from OLE import CStructOLE
        
        l_SSector = []
        l_SSector_Name = []
        l_SSector_Data = []
        
        try :
            StructOLE = CStructOLE()
            
            n_Cnt = 0
            while n_Cnt < dl_OLEDirectory.__len__() :
                #    dl_OLEDirectory[n_Cnt][0]         Short-Sector Name ( MultiByte )
                #    dl_OLEDirectory[n_Cnt][2]         Short-Sector Type
                #    dl_OLEDirectory[n_Cnt][11]        Short-Sector Start SecID
                #    dl_OLEDirectory[n_Cnt][12]        Short-Sector Size
                if dl_OLEDirectory[n_Cnt][2] != 5 and dl_OLEDirectory[n_Cnt][12] < 0x1000 :
                    s_SSector = StructOLE.fnStructOLESector(s_RootEntry, t_SSATList, dl_OLEDirectory[n_Cnt][11], SIZE_OF_SHORT_SECTOR)
                    if s_SSector == None :
                        print "\t\t\t[-] Failure - StructOLESector() for %s in HWP" % (CBuffer.fnExtractAlphaNumber(dl_OLEDirectory[n_Cnt][0]))
                        return False
                    
                    l_SSector_Name.append( CBuffer.fnExtractAlphaNumber(dl_OLEDirectory[n_Cnt][0]) )
                    l_SSector_Data.append( s_SSector )
                    
                n_Cnt += 1
        
        except :
            print format_exc()
            return None
        
        l_SSector.append( l_SSector_Name )
        l_SSector.append( l_SSector_Data )
        
        return l_SSector



class CMappedHWP():
    
    #    s_FileHeader                        [IN]                    File Buffer for HWP FileHeader ( per 0x200  )
    #    t_HWPHeader                         [OUT]                   Mapped HWP Header
    def fnMappedHWPHeader(self, s_FileHeader):
        
        try :
            
            t_HWPHeader_Name = namedtuple('HWPHeader', RULE_HWPHEADER_NAME)
            t_HWPHeader = t_HWPHeader_Name._make( unpack(RULE_HWPHEADER_PATTERN, s_FileHeader) )
            if t_HWPHeader == () :
                return None
        
        except error :
            print format_exc()
            print "\t\t\t0x%08X" % s_FileHeader.__len__()
            return None
        
        except : 
            print format_exc()
            return None
            
        return t_HWPHeader

        
class CExploitHWP():
    
    #    l_SSector                            [IN]                    Short-Sectors
    #            l_SSector[0]                                         Short-Sector's Name
    #            l_SSector[1]                                         Short-Sectos's Data
    #    BOOL Type                            [OUT]                   True / False
    def fnVulnerability(self, l_SSector):
        
        try :
            
            s_SSecName = l_SSector[0]               # Short-Sector Name 
            s_SSecData = l_SSector[1]               # Short-Sector Data               
            
            # Type 1 : Size Overflow - Just "Section"
            for Index in range( s_SSecName.__len__() ) :
                if s_SSecName[Index].find( "Section" ) == 0 :
                    B_Ret, n_TagID = self.fnExploitOverflow(s_SSecName[Index], s_SSecData[Index])
                    if B_Ret == False :
                        print "\t\t\t[-] Failure - ExploitOverflow( %s )" % s_SSecName[Index]
                        return False
                    
                    if n_TagID != None :
                        print "\t\t\t[*] Vulnerability : %s - 0x%02X [ %s ]" % (s_SSecName[Index], n_TagID, HWPTAG[n_TagID])
                        continue
                
            # Type 2 : Embedded PE
            for Index in range( s_SSecName.__len__() ) :
                if self.fnExploitEmbedded(s_SSecData[Index]) == "PE" : 
                    print "\t\t\t[*] Vulnerability : %s - Embedded PE" % (s_SSecName[Index])
                    continue 
            
        except :
            print format_exc()
            return None
        
        return True
        
        
    
    #    s_SSecName                           [IN]                    Short-Sector's Name
    #    s_SSecData                           [IN]                    Short-Sector's Data
    #    BOOL Type                            [OUT]                   True / False
    def fnExploitOverflow(self, s_SSecName, s_SSecData):
        
        try :
            
            try :
                s_Decrypt = decompress( s_SSecData, -15 )
            except :
                s_Decrypt = s_SSecData
            
            n_Position = 0
            while n_Position < s_Decrypt.__len__() :
                n_Record = CBuffer.fnReadDword(s_Decrypt, n_Position)
                if n_Record == None :
                    print "\t\t\t[-] Error - fnReadDword(%s, 0x%08X, 0x%08X)" % (s_SSecName, n_Position, n_Record)
                    return False, None
            
                n_Position += 4
                
                n_TagID = n_Record & 0b1111111111
#                n_Level = ( n_Record >> 10 ) & 0b1111111111
                n_Size = ( n_Record >> 20 ) & 0b111111111111
                if n_Size == 0xfff :
                    n_Size = CBuffer.fnReadDword(s_Decrypt, n_Position)
                    n_Position += 4
                
                n_Position += n_Size
        
            if n_Position == s_Decrypt.__len__() :
                return True, None
        
        except : 
            print format_exc()
            return False, None
        
        return True, n_TagID
        
    
    #    s_SSecData                           [IN]                    Short-Sector's Data 
    #    "PE" or None                         [OUT]                       
    def fnExploitEmbedded(self, s_SSecData):
        
        try :
            
            PE = CPE()
            return PE.Check(s_SSecData)
        
        except : 
            print format_exc()
            return None
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        