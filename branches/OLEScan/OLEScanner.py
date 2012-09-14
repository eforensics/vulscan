# -*- coding:utf-8 -*-

# import Public Module
import os, sys, optparse, traceback, binascii


# import private module
from ComFunc import *



class OLEStruct():
    def OLEHeader(self, File):
        Header = {}
        try : 
            Header = self.DumpHeader(File)
            if Header == {} :
                File['logbuf'] += "\n\t[Failure] DumpHeader()"
                return False
        except :
            print traceback.format_exc()
            return False
        
        File["Header"] = Header
        # referred MSAT
        File["MSAT"] = Header["MSAT"]
        File["SATNumber"] = Header["NumSAT"]
        # referred SAT  
        File["DirSecID"] = Header["DirSecID"]
        # referred Short-SAT
        File["sSATSecID"] = Header["sSATSecID"] 
        File["sSATNumber"] = Header["NumsSAT"]
        
        return True
     
    
    def OLETable(self, File):
        # DumpTable SAT
        SAT = self.DumpTableSAT(File)
        if SAT == [] :
            File['logbuf'] += "\n\t[Failure] DumpTable for SAT"
            return False
        else : 
            File["SAT"] = SAT       
        
        # DumpTable SSAT
        SSAT = self.DumpTableSSAT(File)
        if SSAT == [] :
            File['logbuf'] += "\n\t[Failure] DumpTable for SSAT"
            return False
        else :    
            File["SSAT"] = SSAT  
        
        return True
    
    
    def OLEStorasge(self, File):
        # Dump Storage
        Storage = self.DumpStorage(File)       
        if Storage == "" :
            File['logbuf'] += "\n\t[Failure] DumpStorage for Directory"
            return False
        else :
            File["Storage"] = Storage 
        
        return True
     
    
    def OLEStream(self, File):
        # Directory Dumping
        if not self.OLEDirectory(File) :
            File['logbuf'] += "\n\t[Failure] OLEstruct.OLEDirectory()"
            return False
        
        # Stream Dumping
        if not self.DumpStream(File) :
            File['logbuf'] += "\n\t[Failure] OLESturct.DumpStream()"
            return False        
        
        return True
    
    
    def OLEDirectory(self, File):
        Directory = {}
        Storage = File["Storage"]
        DirCnt = len(Storage) / szData["Directory"]
        
        try :
            Cnt = 0 
            while Cnt < DirCnt :
                DirData = BufferControl.Read(Storage, Cnt * szData["Directory"], szData["Directory"])
                Directory = self.DumpDirectory(DirData)
                if Directory == {} :
                    File['logbuf'] += "\n\t[Failure] OLESturct.DumpDirectory() - %x" % Cnt
                    return False
                
                DirName = BufferControl.ExtractAlphaNumber( Directory["EntryName"] )
                if DirName == "" :
                    DirName = "Empty" + str(Cnt)                                                             
                File[DirName] = Directory
                Cnt += 1
                
                if Directory["EntryType"] == '05' or Directory["szData"] > 0x1000 :
                    ReferredSAT.append( DirName )
                elif Directory["EntryType"] != '05' and Directory["szData"] <= 0x1000 :
                    ReferredSSAT.append( DirName )
                else :
                    File['logbuf'] += "\n\t[Failure] Failed Separation by Referred Type ( %s )" % Directory["EntryName"]
                    return False
                
#                File['logbuf'] += "\n%s" % DirName
#                self.PrintDirectoryLog(Directory)
            
            if Cnt == DirCnt :
                File["DirCnt"] = Cnt
            else :
                File['logbuf'] += "\n\t[Failure] OLEStruct.OLEDirectory() ( Cnt : %x, DirCnt : %x )" % (Cnt, DirCnt)            
                return False
        except :
            print traceback.format_exc()
            
        return True


    def PrintDirectoryLog(self, Directory):
        EntryName = BufferControl.Read(Directory["EntryName"], 0, Directory["szName"])
        File['logbuf'] += "\n%s" % EntryName
        if Directory["EntryType"] == '01' :
            File['logbuf'] += "\tStorage"                      
        elif Directory["EntryType"] == '02' : 
            File['logbuf'] += "\tStream"
        elif Directory["EntryType"] == '05' :
            File['logbuf'] += "\tRoot" 
        else :
            File['logbuf'] += "\tEmpty"
                
        File['logbuf'] += "\t0x%08X" % Directory["LeftChild"]
        File['logbuf'] += "\t0x%08x" % Directory["RightChild"] 
        File['logbuf'] += "\t0x%08x" % Directory["SubNode"] 
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
        
        return True


    def OLEStreamScanning(self, File):
        File['logbuf'] += "\n\tOLEStruct.OLEStreamScanning() - Not Yet"
        return True    
    
    
    def DumpHeader(self, File):
        pBuf = File["pBuf"]
        
        index = 0
        Position = 0
        MSAT = []
        Header = {}
        OutHeader = {}
        try : 
            for index in range( len(szHeader) ) :
                if index == ( len(szHeader) - 1 ) :
                    MSAT = self.DumpTable(pBuf, Position, 0x1B4)
                    if MSAT == [] :
                        break
                    
                    Header[ mHeader[index] ] = MSAT
                    OutHeader = Header
                    break            
                
                if szHeader[index] == 2 :
                    Header[ mHeader[index] ] = BufferControl.ReadWord(pBuf, Position)
                elif szHeader[index] == 4 :
                    Header[ mHeader[index] ] = BufferControl.ReadDword(pBuf, Position)
                else :
                    Header[ mHeader[index] ] = binascii.b2a_hex( BufferControl.Read(pBuf, Position, szHeader[index]) )
            
                Position += szHeader[index]
        except : 
            print traceback.format_exc()
        
        # Exception Checking
        if Header["NumSAT"] > 109 :   
            File['logbuf'] += "\n\t[Failure] Exist Extra SAT!!!"
            OutHeader = {}        
        
        return OutHeader
    
    
    def DumpStorage(self, File):
        Storage = ""
        pBuf = File["pBuf"]
        SAT = File["SAT"]
        SecID = File["DirSecID"]
        
        try :
            Sector = ""
            
            while True :
                Sector += BufferControl.ReadSectorByBuffer(pBuf, SecID, szData["Sector"])
                SecID = SAT[SecID]
                if SecID == 0xfffffffe :
                    break
            
            if Sector != "" :
                Storage = Sector
        except :
            print traceback.format_exc()        
        
        return Storage
    
    
    def DumpTable(self, pBuf, Position, Size):
        table = []
        StartPosition = Position
        
        try : 
            while Position < StartPosition + int(Size) : 
                table.append( BufferControl.ReadDword(pBuf, Position) )
                Position += 4
        except :
            print traceback.format_exc()
        
        return table
    
    
    def DumpTableSAT(self, File):
        MSAT = File["MSAT"]
        pBuf = File["pBuf"]
        
        SAT = []   
        
        try :
            Sector = ""
            count = 0
            index = 0

            for index in range( len(MSAT) -1 ) :
                if MSAT[index] == 0xffffffffL or MSAT[index] == 0xfffffffeL :
                    break
                
                Sector += BufferControl.ReadSectorByBuffer(pBuf, MSAT[index], szData["Sector"])
                count += 1
            
            SAT = self.DumpTable(Sector, 0, szData["Sector"] * count)
        except :
            print traceback.format_exc()
        
        return SAT
        
        
    def DumpTableSSAT(self, File):
        number = File["sSATNumber"]
        SecID = File["sSATSecID"]
        pBuf = File["pBuf"]
        
        SSAT = []
        
        try :
            Sector = ""
            if number > 1 : 
                Sector = self.DumpTableSAT(File)
            else :
                Sector += BufferControl.ReadSectorByBuffer(pBuf, SecID, szData["Sector"])
            
            SSAT = self.DumpTable(Sector, 0, szData["Sector"] * number)
        except :
            print traceback.format_exc()
            
        return SSAT


    def DumpDirectory(self, pBuf):
        Position = 0
        index = 0
        Directory = {}
        
        try : 
            for index in range( len(szDirEntry) ) :
                if szDirEntry[index] == 2 :
                    Directory[ mDirEntry[index] ] = BufferControl.ReadWord(pBuf, Position)
                elif szDirEntry[index] == 4 :
                    Directory[ mDirEntry[index] ] = BufferControl.ReadDword(pBuf, Position)
                elif szDirEntry[index] == 64 :
                    Directory[ mDirEntry[index] ] = BufferControl.Read(pBuf, Position, szDirEntry[index])
                else :
                    Directory[ mDirEntry[index] ] = binascii.b2a_hex( BufferControl.Read(pBuf, Position, szDirEntry[index]))
        
                Position += szDirEntry[index]
        except :
            print traceback.format_exc()
        
        return Directory
        
    
    def DumpStream(self, File):
        for stream in ReferredSAT :
            if not self.DumpStreamSAT(File, stream) :
                File['logbuf'] += "\n\t[Failure] OLEStruct.DumpStreamSAT() - %s" % stream
                return False
        
        for stream in ReferredSSAT :
            if not self.DumpStreamSSAT(File, stream) :
                File['logbuf'] += "\n\t[Failure] OLEStruct.DumpStreamSSAT() - %s" % stream
                return False

        return True
    

    def DumpStreamSAT(self, File, DirName):
        Stream = ""
        Directory = File[DirName]
        SecID = Directory["SecID"]
        
        try :
            Sector = ""
            SAT = File["SAT"]
            pBuf = File["pBuf"]
            
            while True :
                Sector += BufferControl.ReadSectorByBuffer(pBuf, SecID, szData["Sector"])
                SecID = SAT[SecID]
                if SecID == 0xfffffffe :
                    break
        
            if Sector != "" :
                Stream = Sector
            else :
                return False
        
            fname = "%s_%s.dump" %(File["fname"], DirName)
            FileControl.WriteFile(fname, Stream)
            
            EntryName = BufferControl.ExtractAlphaNumber( Directory["EntryName"] )
            if EntryName == "RootEntry" :
                File["pRoot"] = Stream
            
        except :
            print traceback.format_exc()
        
        return True


    def DumpStreamSSAT(self, File, DirName):
        Stream = ""
        Directory = File[DirName]
        SecID = Directory["SecID"]
        
        try :
            Sector = ""
            SSAT = File["SSAT"]
            pBuf = File["pRoot"]
            
            while True :
                Sector += BufferControl.ReadSectorByBuffer(pBuf, SecID, szData["ShortSector"])
                SecID = SSAT[SecID]
                if SecID == 0xfffffffe :
                    break
        
            if Sector != "" :
                Stream = Sector
            else :
                return False
        
            fname = "%s_%s.dump" %(File["fname"], DirName)
            FileControl.WriteFile(fname, Stream)
            
        except :
            print traceback.format_exc()
        
        return True



ReferredSAT = []
ReferredSSAT = []

szData = {"MSAT":0x1B4, "SAT":0x1FF, "SSAT":0x1FF, "Directory":0x80, "Sector":0x200, "ShortSector":0x40}

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

# Table SecID's Description
#            SECID_FREE          = -1 # 0xFFFFFFFF - Free Sector may exist in the file, but is not part of any stream
#            SECID_END_OF_CHAIN  = -2 # 0xFFFFFFFE - Trailing SecID in a SecID chain
#            SECID_SAT           = -3 # 0xFFFFFFFD - Sector is used by the sector allocation table
#            SECID_MSAT          = -4 # 0xFFFFFFFC - Sector is used by the master sector allocation table




class OLEScan():
    
    @classmethod
    def Check(cls, fname):
        try :
            with open( fname, 'rb' ) as checkfile :
                pBuf = checkfile.read()
                
                if BufferControl.ReadDword(pBuf, 0) == 0xe011cfd0L and BufferControl.ReadDword(pBuf, 0x4) == 0xe11ab1a1L :
                    return True
                else : 
                    return False
        except :
            print traceback.format_exc()
        
        
    @classmethod
    def Scan(cls, File):
        File["pfile"] = FileControl.OpenFileByBinary(fname)
        File["pBuf"] = File["pfile"].read()
        
        
        File['logbuf'] += "\n[*] Scanning"
        try :
            OLEStructure = OLEStruct()
            
            if OLEStructure.OLEHeader(File) :
                File['logbuf'] += "\n\t[*] Mapped OLEHeader"
            else :
                File['logbuf'] += "\n\t[Failure] Mapped OLEHeader"
                return
            
            
            if OLEStructure.OLETable(File) :
                File['logbuf'] += "\n\t[*] Mapped OLETable"
            else :
                File['logbuf'] += "\n\t[Failure] Mapped OLETable"
                return
                
            
            if OLEStructure.OLEStorasge(File) :
                File['logbuf'] += "\n\t[*] Mapped OLEStorage"
            else :
                File['logbuf'] += "\n\t[Failure] Mapped OLEStorage"
                return
                
            
            if OLEStructure.OLEStream(File) :
                File['logbuf'] += "\n\t[*] Mapped OLEStream"
            else :
                File['logbuf'] += "\n\t[Failure] Mapped OLEStream"
                return            
            
            
            if OLEStructure.OLEStreamScanning(File) :
                File['logbuf'] += "\n\t[*] Scanning OLEStream"
            else :
                File['logbuf'] += "\n\t[Failure] Scanning OLEStream"
                return
            
        except :
            print traceback.format_exc()
        
        
        if File["pfile"] :
            File["pfile"].close()
        
        
    


class OLEInit():
    def __init__(self, File):
        # Init Input Options.......
        File['fpath'] = ""
        File['dpath'] = ""
        File['log'] = ""
        File['logbuf'] = ""
        
        # Init OLEStruct..........
        File["Header"] = {}
        File["Storage"] = ""
        File["MSAT"] = []
        File["SAT"] = []
        File["SSAT"] = []
        
        # Init File Information.......
        # File["fname"]
        # File["pfile"]
        # File["pBuf"]
        
        return


    def GetOption(self) : 
        Parser = optparse.OptionParser(usage='usage: %prog [-f|-d] file or folder\n')
        Parser.add_option('-f', '--file', help='< file name >')
        Parser.add_option('-d', '--directory', help='< directory >')
        Parser.add_option('--log', help='< file name >')
        (options, args) = Parser.parse_args()
        
        if len( sys.argv ) < 2 :
            Parser.print_help()
            exit(0)
    
        return options
    

    def SetOption(self, File, Options):
        if Options.file :
            File['fpath'] = Options.file
        
        if Options.directory :
            File['dpath'] = Options.directory
            
        if Options.log :
            File['log'] = Options.log
        
        return
        
    
    def PrintLog(self, File):
        if File['log'] == "" :
            print File['logbuf']
        else :
            FileControl.WriteFile( File['log'], File['logbuf'] )





if __name__ == '__main__' :
    File = {}
    
    OLE = OLEInit(File)
    
    File['logbuf'] += "[*] OLEScanner.py"
    Options = OLE.GetOption()
    try :
        OLE.SetOption(File, Options)
        
        flist = []
        if File['fpath'] != "" :
            flist.append( File['fpath'] )
        elif File['dpath'] != "" :
            flist = os.listdir( File['dpath'] )
            os.chdir( File['dpath'] )
        else :
            File['logbuf'] += "\n[Failure] Get File List for Scanning"
            exit(-1) 
    
        
        for fname in flist :
            if OLEScan.Check(fname) :
                File['logbuf'] +="\n[*] Check File : %s..........OLE" % fname
                File['fname'] = fname
            else :
                File['logbuf'] +="\n[*] Check File : %s..........NOT OLE" % fname
                continue
            
            File["fname"] = fname
            OLEScan.Scan(File)
    
    except :
        print traceback.format_exc()
        
    OLE.PrintLog(File)
    
    exit(0)
    
            
        
    











