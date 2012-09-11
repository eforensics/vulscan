# -*- coding:utf-8 -*-

# import Public Module
import os, sys, optparse, traceback, binascii


# import private module
from CtrlFile import CTRL



class OLEStruct():
    def __init__(self):
        return
    

    @classmethod
    def DumpOLEHeader(cls, File, fname, fbuf):
        Header = {}
        try : 
            Header = OLEStruct.Header(fbuf)
            if Header == {} : 
                File['logbuf'] += "\n[ERR] Failed - OLEStruct.Header( %s )" % fname
                return False    
        except :
            File['logbuf'] += "\n[ERR] OLEStruct.Header( %s )" % fname + traceback.format_exc()        
        
        File['Header'] = Header
        File['MSAT'] = Header['MSAT']
        return True
    

    @classmethod
    def DumpSAT(cls, File, fname, fp, MSAT, NumSAT):
        SAT = ""
        try :
            SAT = OLEStruct.TraceSAT(fp, MSAT, NumSAT)
            if SAT == "" :
                File['logbuf'] += "\n[ERR] Failed - OLEStruct.TraceSAT( %s )" % fname
                return False    
        except :
            File['logbuf'] += "\n[ERR] OLEStruct.TraceSAT( %s )" % fname + traceback.format_exc()
        
        File['SAT'] = SAT
                
        fname_SAT = fname + "_SAT.dump"
        with open( fname_SAT, 'wb' ) as SAT_write :
            SAT_write.write(SAT)

        return True
    
    
    @classmethod
    def DumpSSAT(cls, File, fname, fp, SecID, NumSSAT):
        SSAT = ""
        try :
            SSAT = OLEStruct.TraceSSAT(fp, File['MSAT'], SecID, NumSSAT)
            if SSAT == "" : 
                File['logbuf'] += "\n[ERR] Failed - OLEStruct,TraceSSAT( %s )" % fname
                return False
        except :
            File['logbuf'] += "\n[ERR] OLEStruct.TraceSSAT( %s )" % fname + traceback.format_exc()
            
        File['SSAT'] = SSAT
        
        fname_SSAT = fname + "_SSAT.dump"
        with open( fname_SSAT, 'wb' ) as SSAT_write :
            SSAT_write.write( SSAT )
            
        return True
        
    
    
    @classmethod
    def DumpStorage(cls, File, fname, fp):
        Storage = ""
        SAT = File["SAT"]
        
        StartIndex = File["Header"]["DirSecID"]
        SAT_list = CTRL.Bin2List( SAT, 4 )        
         
        try :
            Storage = OLEStruct.TraceTable(fp, SAT_list, StartIndex)
            if Storage == "" :
                File['logbuf'] += "\n[ERR] Failed - OLEStruct.TraceTable( %s )" % fname
                return False
        except :
            File['logbuf'] += "\n[ERR] OLEStruct.TraceTable( %s )" % fname + traceback.format_exc()
        
        File['Storage'] = Storage
        
        fname_Storage = fname + "_Storage.dump"
        with open( fname_Storage, 'wb' ) as Storage_write :
            Storage_write.write( Storage )

        return True


    @classmethod
    def DumpStream(cls, File, fbuf, table):
        pass
    
    
    @classmethod
    def FindFamily(cls, File, fbuf, Storage):
#        print "OLEStruct.FindFamily()"
        OLEFamily = ""
        Directory = {}
        offset = 0
        # Mapped Directory
        while offset < len(Storage) :
            DirData = CTRL.Read(Storage, offset, szDirectory)
            try : 
                Directory = OLEStruct.Directory(DirData)  
            except :
                File['logbuf'] += "\n[ERR] OLEStruct.Directory()" + traceback.format_exc()
            
            offset +=  szDirectory
            if offset == len(Storage) :
                break                    
            
##           Check OLE Family 
##           |HWPSummaryInformation             
##           05 00 48 00 77 00 70 00 53 00 75 00 6D 00 6D 00 61 00 72 00 79 00 49 00 6E 00 66 00 6F 00 72 00 6D 00 61 00 74 00 69 00 6F 00 6E 00 00 00
#            if "HwpSummaryInformation".encode( "utf-16le" ) in Directory["EntryName"] :
#                OLEFamily = "HWP"
#
##           WordDocument
##           57 00 6F 00 72 00 64 00 44 00 6F 00 63 00 75 00 6D 00 65 00 6E 00 74 00 00 00
#            if "WordDocument".encode( "utf-16le" ) in Directory["EntryName"] :
#                OLEFamily = "DOC"
#
##           Workbook
##           57 00 6F 00 72 00 6B 00 62 00 6F 00 6F 00 6B 00 00 00
#            if "Wordbook".encode( "utf-16le" ) in Directory["EntryName"] :
#                OLEFamily = "XLS"
#
##           PowerPoint Document
##           50 00 6F 00 77 00 65 00 72 00 50 00 6F 00 69 00 6E 00 74 00 20 00 44 00 6F 00 63 00 75 00 6D 00 65 00 6E 00 74 00 00 00
#            if "PowerPoint Document".encode( "utf-16le" ) in Directory["EntryName"] :
#                OLEFamily = "PPT"
            
            
            #######################################################            
            #             Directory Log......
            #######################################################
            EntryName = CTRL.Read(Directory["EntryName"], 0, Directory["szName"])
            File['logbuf'] += "\n[%d]" %(offset/szDirectory -1) + EntryName
#            File['logbuf'] += "\n[%d]" %(offset/szDirectory -1) + Directory["EntryName"]
            if Directory["EntryType"] == '01' :
                File['logbuf'] += "\tStorage"          
            elif Directory["EntryType"] == '02' : 
                File['logbuf'] += "\tStream"
            elif Directory["EntryType"] == '05' :
                File['logbuf'] += "\tRoot" 
            else :
                File['logbuf'] += "\tEmpty"
                
            File['logbuf'] += "\t" + hex( Directory["LeftChild"] )
            File['logbuf'] += "\t" + hex( Directory["RightChild"] )
            File['logbuf'] += "\t" + hex( Directory["SubNode"] )
            File['logbuf'] += "\t" + hex( Directory["SecID"] )
            File['logbuf'] += "\t" + hex( Directory["szData"] )
                
            if Directory["EntryType"] == '05' :
                File['logbuf'] += "\tReferred SAT"
            elif Directory["EntryType"] != '05' and Directory["szData"] > 0x1000 :
                File['logbuf'] += "\tReferred SAT"
            elif Directory["EntryType"] != '05' and Directory["szData"] <= 0x1000 :
                File['logbuf'] += "\tReferred Short-SAT"
            else :
                File['logbuf'] += "Error"            
        
        return OLEFamily

    
    
    @classmethod
    def Header(cls, fbuf):
#        print "OLEStruct.Header()"
        
        index = 0
        offset = 0
        MSAT = []
        Header = {}
        OutHeader = {}
        for index in range( len(szHeader) ) :
            if index == ( len(szHeader) - 1 ) :
                MSAT = OLEStruct.MapTable(fbuf, offset, szHeader[index])
                if MSAT == [] :
                    return OutHeader
                
                Header[ mHeader[index] ] = MSAT
                OutHeader = Header
                break
            
            if szHeader[index] == 2 : 
                Header[ mHeader[index] ] = CTRL.Word(fbuf, offset)
            elif szHeader[index] == 4 :
                Header[ mHeader[index] ] = CTRL.Dword(fbuf, offset)
            else :
                Header[ mHeader[index] ] = binascii.b2a_hex( CTRL.Read(fbuf, offset, szHeader[index]) )
                            
            offset += szHeader[index]        
        
        # Check Header Data
        if Header["NumSAT"] > 109 :
            print "Please refer to the Extra SAT!"
        
        return OutHeader


    @classmethod
    def Directory(cls, fbuf):
        offset = 0
        Directory = {}
        for index in range( len(szDirEntry) ) :
            if szDirEntry[index] == 2 :
                Directory[ mDirEntry[index] ] = CTRL.Word(fbuf, offset)
            elif szDirEntry[index] == 4 :
                Directory[ mDirEntry[index] ] = CTRL.Dword(fbuf, offset) 
            elif szDirEntry[index] == 64 :
                Directory[ mDirEntry[index] ] = CTRL.Read(fbuf, offset, szDirEntry[index]) 
            else :
                Directory[ mDirEntry[index] ] = binascii.b2a_hex( CTRL.Read(fbuf, offset, szDirEntry[index]) )                
            
            offset += szDirEntry[index]
        
        return Directory


    # From DumpData to List for table
    @classmethod
    def MapTable(cls, fbuf, offset, size):
#        print "OLEStruct.MapTable()"
        
        index = 0
        table = []
        startoffset = offset
        
        while offset < startoffset + size :
            try :
                table.append (CTRL.Dword(fbuf, offset) )            
            except :
                print traceback.format_exc()

            offset += 4
            index += 1

        return table


    # trace for SAT in MSAT
    @classmethod
    def TraceSAT(cls, fp, MSAT, NumSAT):
#        print "OLEStruct.TraceSAT()"        
        
        Sector = ""
        SAT = ""
        
        for i in range(len(MSAT) -1) :
            if MSAT[i] == 0xffffffff :
                continue
            
            try :
                Sector += CTRL.ReadSector(fp, MSAT[i])
                if ( i + 1 ) == NumSAT :
                    SAT = Sector
                    break
            except :
                print traceback.format_exc()        
            
        return SAT
    
    
    @classmethod
    def TraceSSAT(cls, fp, MSAT, SecID, NumSSAT):
#        print "OLEStruct.TraceSSAT()"
        SSAT = ""
        Sector = ""

        try :
            if NumSSAT > 1 : # TraceSAT
                Sector = OLEStruct.TraceSAT(fp, MSAT, NumSSAT)
            else :
                Sector += CTRL.ReadSector(fp, SecID)
        except :
            print traceback.format_exc()
        
        SSAT = Sector
        
        return SSAT

    
    # trace for Sector in SAT/SSAT
    @classmethod
    def TraceTable(cls, fp, Table, StartIndex):
        Sector = ""
        
        SecID = StartIndex
        
        while True :
            Sector += CTRL.ReadSector(fp, SecID)
            SecID = Table[SecID]
            
#            SECID_FREE          = -1 # 0xFFFFFFFF - Free Sector may exist in the file, but is not part of any stream
#            SECID_END_OF_CHAIN  = -2 # 0xFFFFFFFE - Trailing SecID in a SecID chain
#            SECID_SAT           = -3 # 0xFFFFFFFD - Sector is used by the sector allocation table
#            SECID_MSAT          = -4 # 0xFFFFFFFC - Sector is used by the master sector allocation table
            if SecID == 0xFFFFFFFF or SecID == 0xFFFFFFFE :
                return Sector



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


szDirectory = 0x80
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


#-------------------------------------------------------------------------------------------------------------------------

class OLE():
    def __init__(self, File):
        # Input Options.....
        File['fpath'] = ""
        File['dpath'] = ""
        # File['scan'] = ""
        # File['delete'] = ""
        # File['vtotal'] = ""
        File['log'] = ""
        File['logbuf'] = ""
        # File['search'] = ""
        
        # Proceduring OLE
        File['Header'] = {}
        File['MSAT'] = ""
        File['SAT'] = ""
        File['SSAT'] = ""
        File['Storage'] = "" 
        return
    
    
    @classmethod
    def GetOpt(cls, File):
        # print "OLE.GetOpt()"
        
        Parser = optparse.OptionParser(usage='usage: %prog [-f|-d] file or folder\n')
        Parser.add_option('-f', '--file', help='< file name >')
        Parser.add_option('-d', '--directory', help='< directory >')
        # Etc.....
        Parser.add_option('--log', help='< file name >')
        (options, args) = Parser.parse_args()
        
        if len( sys.argv ) < 2 :
            Parser.print_help()
            exit(0)
    
        return options
        
    
    @classmethod
    def SetOpt(cls, File, options):
        # print "OLE.SetOpt()"
        
        if options.file :
            File['fpath'] = options.file 
        
        if options.directory :
            File['dpath'] = options.directory
            os.chdir( options.directory )
        
        if options.log :        # Arg : File Name ( ABS or Relative )
            File['log'] = options.log
        
        return
    
    
    @classmethod
    def PrintLog(cls, File):
        # print "OLE.PrintLog()"
        
        if File['log'] == "" :
            print File['logbuf']
        else :
            with open( File['log'], 'w' ) as logfile :
                logfile.write( File['logbuf'] )
                
        return
        
    
    # Check OLE Signature : 0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1
    # return : True / False
    @classmethod
    def Check(cls, fname): 
        # print "OLE.Check()" 
             
        with open( fname, 'rb' ) as checkfile :
            fp = checkfile.read()
    
            # Little-Endian
            if CTRL.Dword(fp, 0x0) == 0xe011cfd0L and CTRL.Dword(fp, 0x4) == 0xe11ab1a1L :
                return True

        return False
    
        
    # Check OLE's Family : Office Family ( DOC, PPT, XLS ) | Korean Editor Family ( HWP )
    # return : DOC, PPT, XLS, HWP
    @classmethod
    def CheckFamily(cls, File, fname):
        # print "OLE.Scan.DeepCheck()"
        fp = open( fname, 'rb' )    
        fbuf = fp.read()
        
        
        # Mapped OLE Header
        if OLEStruct.DumpOLEHeader(File, fname, fbuf) :
            File['logbuf'] += "\n[*] Mapped OLEHeader"
        else :
            File['logbuf'] += "\n[ERR] Failed - OLEStruct.DumpOLEHeader()" + traceback.format_exc()
               
        
        # Dumped SAT ( Output : fname_SAT.dump )
        if OLEStruct.DumpSAT(File, fname, fp, File["MSAT"], File["Header"]["NumSAT"]) :
            File['logbuf'] += "\n[*] Dumped SAT"
        else :
            File['logbuf'] += "\n[ERR] Failed - OLEStruct.DumpSAT()" + traceback.format_exc()
        
        
        # Dumped SSAT ( Output : fname_SSAT.dump )
        if OLEStruct.DumpSSAT(File, fname, fp, File["Header"]["sSATSecID"], File["Header"]["NumsSAT"]) :
            File['logbuf'] += "\n[*] Dumped SSAT"
        else:
            File['logbuf'] += "\n[ERR] Failed - OLEStruct.DumpSSAT()" + traceback.format_exc()
        
        
        # Dumped Storage ( Output : fname_Storage.dump )
        if OLEStruct.DumpStorage(File, fname, fp) :
            File['logbuf'] += "\n[*] Dumped Storage"
        else :
            File['logbuf'] += "\n[ERR] Failed - OLEStruct.DumpStorage()" + traceback.format_exc()
        
        
        # Find OLE Family
        File['logbuf'] += "\n\n" + "[*] View Directory"
        File['logbuf'] += "\n" + "==========" * 12
        File['logbuf'] += "\n DirName\t\t\t\t\tType\tLeft\tRight\tSub\tSecID\tSize\tTable"
        File['logbuf'] += "\n" + "==========" * 12
            
        OLEFamily = ""
        OLEFamily = OLEStruct.FindFamily(File, fbuf, File['Storage'])
        if  OLEFamily == "" :
            File['logbuf'] += "\n[ERR] Failed - OLEStruct.FindFamily()"
        
        return OLEFamily    
    
    
    @classmethod
    def Scan(cls, File, fname):
        # print "OLE.Scan()"
        
        try : 
            OLEFamily = OLE.CheckFamily(File, fname)
        except :
            File['logbuf'] += "\n[ERR] OLE.CheckFamily(%s)" % fname + traceback.format_exc()
            return        
        
        if OLEFamily == "DOC" or OLEFamily == "XLS" or OLEFamily == "PPT" :
            try :
                OLE.ScanOffice(File, fname)
            except :
                File['logbuf'] += "\n[ERR] OLE.ScanOffice(%s)" % fname + traceback.format_exc()
        elif OLEFamily == "HWP" :
            try :
                OLE.ScanHWP(File, fname)
            except :
                File['logbuf'] += "\n[ERR] OLE.ScanHWP(%s)" % fname + traceback.format_exc()
        else :
            File['logbuf'] += "\n[*] Do not Support OLE File Format (%s : %s )" % (OLEFamily, fname)     
        
        return    
    
    # Target : doc, xls, ppt ( do not support docX ,pptx )
    @classmethod
    def ScanOffice(cls, File, fname):
        print "OLE.Scan.ScanOffice()"
        pass
    
    
    # Target : hwp
    @ classmethod
    def ScanHWP(cls, File, fname):
        print "OLE.Scan.ScanHWP()"
        pass
    
    

if __name__ == '__main__' :
    File = {}
    ole = OLE(File)
    
    File['logbuf'] += "[*] OLEParser"
    
    try :
        options = ole.GetOpt(File)
    except :
        File['logbuf'] += "\n[ERR] GetOpt()\n" + traceback.format_exc()
        exit(-1)
    
    try :
        ole.SetOpt(File, options)
    except :
        File['logbuf'] += "\n[ERR] SetOpt()\n" + traceback.format_exc()
        exit(-1)
    
    
    flist = []
    if File['fpath'] != "" and File['dpath'] == "" :
        flist.append( File['fpath'] )
#        File['logbuf'] += File['fpath']
    elif File['fpath'] == "" and File['dpath'] != "" :
        flist = os.listdir( File['dpath'] )
#        File['logbuf'] += File['dpath']
    else :
        File['logbuf'] += "\n[ERR] Failed - Get File List\n" + traceback.format_exc()
            
    
    for f in flist :
        File['logbuf'] += "\n\n[*] Target : "
        File['logbuf'] += f
        try :
            if ole.Check( f ) :
                File['logbuf'] += "..........OLE"
            else :
                File['logbuf'] += "..........NOT OLE" 
                continue
        except :
            File['logbuf'] += "\n[ERR] Failed - OLE.Check(%s)\n" % f + traceback.format_exc()
            continue
        
        try :
            ole.Scan(File,  f)
        except :
            File['logbuf'] += "\n[ERR] Failed - OLE.Scan(%s)\n" % f + traceback.format_exc()
            continue
    
    try :
        ole.PrintLog(File)
    except :
        print "\n[ERR] OLE.PrintLog()\n" + traceback.format_exc()
    
    
    exit(0)
    
    
    
    
    
    
    
    