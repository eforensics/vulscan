# -*- coding:utf-8 -*-

# import Public Module
import optparse, traceback, sys, os, ftplib, shutil

# import Private Module
from Common import FileControl
from PDFScanner import PDFScan
from OLEScanner import OLEScan, OLEStruct
from PEScanner import PEScan



class Initialize():
    def GetOptions(self):
        Parser = optparse.OptionParser(usage='usage: %prog [--IP] IP [--ID] LogID [--PW] LogPW [--src] SrcDir [--dst] DstDir\n')
        Parser.add_option('--ip', help='< IP >')
        Parser.add_option('--id', help='< ID >')
        Parser.add_option('--pw', help='< Password >')
        Parser.add_option('--src', help='< Absolutely Source Directory in Server >')
        Parser.add_option('--dst', help='< Absolutely Destination Directory in Client >')
        
        Parser.add_option('-d', '--directory', help='< Delete Directory >')
        
        Parser.add_option('--scan', help='< Format >')
        Parser.add_option('--delete', help='< Extend >')
        
        Parser.add_option('--log', help='< file name >')
        (options, args) = Parser.parse_args()
            
        if len( sys.argv ) < 2 :
            Parser.print_help()
            exit(0)
    
        return options



class FTPServer():
    def ConnectServer(self, FTP, IP, ID, PW):
        try :
            # Connect Network
            print "[+] Connect Server : %s ( ID : %s ).........." % (IP, ID),
            MsgConnect = FTP.connect(IP)
            
            # Failure Connection
            #    - Success Message : 220 (vsFTPd 2.3.5)
            if MsgConnect.find("220") == -1 :
                print "Failure ( Connect : %s )" % MsgConnect
                return False
            
            # Failure Connection
            #    - Success Message : 230 Login successful
            MsgConnect = FTP.login(ID, PW)
            if MsgConnect.find("230") == -1 :
                print "Failure ( Login : %s )" % MsgConnect
                return False
            
            # Success Connection
            print "Success"
            
        except :
            print traceback.format_exc()
            return False
        
        return True
    
    
    def Download(self, FTP, Src, Dst):
        try :
            print "[+] Download......",
            
            if not os.path.isdir( Dst ) :
                os.mkdir( Dst )
            
            os.chdir( Dst )
            SrcList = Src.split("\\")
            
            for nextdir in SrcList :
                MsgFTP = FTP.cwd(nextdir)
                if MsgFTP.find("250") == -1 :
                    print "Failure Change Directory"
                    return False            
            
            if not self.DownloadFile(FTP) :
                return False
            
            print "Success"
            
        except :
            print traceback.format_exc()
            return False
        
        return True
    
    
    def DownloadFile(self, FTP):
        try :
            # Check Target Directory By File Format
            TargetName = ["PDF", "unknown", "MS Compress", "MS Excel Spreadsheet", "MS Word Document", "Office Open XML Document" \
                          "Office Open XML Spreadsheet"]
            ExtDir = []
            
            DirList = FTP.nlst()
            for DirName in DirList : 
                for Target in TargetName :
                    if DirName == Target :
                        ExtDir.append( DirName )
            
            if ExtDir == [] :
                print "Folder is not Found"
                return True
            

            SrcPath = FTP.pwd()
            for extDir in ExtDir :
                # Move Src Directory ( in FTP Server )
                srcpath = SrcPath + "/" + extDir
                MsgFTP = FTP.cwd( srcpath )
                if MsgFTP.find("250") == -1 :
                    print "Failure Change Directory ( %s )" % extDir
                
                # File Download ( From FTP Server To Host PC )
                flist = FTP.nlst()
                for fname in flist :
                    if os.path.splitext( fname )[1] == ".txt" :
                        continue
                    
                    MsgFTP = FTP.retrlines("RETR " + fname, open(fname, 'wb').write)
                    if MsgFTP.find( "226" ) == -1 :
                        print "    %s ( %s )" % ( fname, MsgFTP )
                        continue
            
        except :
            print traceback.format_exc()
            return False
        
        return True


class Action():
    def OptDelete(self, Options):
        try :           
            ext = os.path.splitext( Options.delete )
            # Case 1. test.ext
            if ext[0] != "*" :
                os.remove( Options.delete )
                return True
            
            # Case 2. *
            if ext[0] == "*" and ext[1] == "" :
                os.removedirs( Options.delete )
                return True
            
            # Case 3. *.extend
            flist = os.listdir( os.path.curdir() )
            for fname in flist :
                if os.path.isdir( fname ) :
                    continue
                
                fext = os.path.splitext( fname )    
                if ext[1] == fext[1] :
                    os.remove( fname )
            
            return True
            
        except : 
            print traceback.format_exc()
            return False

    def OptFTP(self, Options) :
        try : 
            IP = Options.ip
            ID = Options.id
            PW = Options.pw
            srcdir = Options.src
            dstdir = Options.dst
            
            if not IP or not ID or not PW or not srcdir or not dstdir :
                print "   [OptErr] Check please FTP Options"
                return False    
            
            # Connection FTP Server
            FTP = ftplib.FTP()
            Server = FTPServer()
            if not Server.ConnectServer(FTP, IP, ID, PW) :
                return False
            
            # Download Files From FTP Server
            if not Server.Download(FTP, srcdir, dstdir) :
                return False
            
            # Connection Termination
            FTP.close()
        
            return True
        except :
            print traceback.format_exc()
            return False


    def Categorizer(self, Options):
        try :    
            # Separate File Format
            FileList = {}
            print "[+] Check Samples File Format"
            if not self.SeparateList(Options, FileList) :
                return {}
            
            # Separate Samples
            FileList["Except"] = []
            
            if FileList["PDF"] != [] :
                self.Separation(Options, FileList["PDF"], "PDF", FileList["Except"]) 
                
            if FileList["OLE"] != [] :
                self.Separation(Options, FileList["OLE"], "OLE", FileList["Except"]) 
                
            if FileList["PE"] != [] :
                self.Separation(Options, FileList["PE"], "PE", FileList["Except"])
            
            if FileList["Unknown"] != [] :
                self.Separation(Options, FileList["Unknown"], "unknown", FileList["Except"]) 

            return FileList
        
        except :
            print traceback.format_exc()
            return {}


    def SeparateList(self, Options, FileList):
        try :
            PDFList = []
            OLEList = []
            PEList = []
            UnknownList = []
            
            if Options.dst :
                dstdir = Options.dst 
            elif Options.directory :
                dstdir = Options.directory 
            else :
                print "Do not Sample's Directory"
                return False
            
            main = Main()
            flist = os.listdir( dstdir )
            for fname in flist :
                if os.path.isdir( fname ) :
                    continue
                
                Format = main.CheckFormat(fname)
                if Format == "PDF" :
                    PDFList.append( fname )
                elif Format == "OLE" :
                    OLEList.append( fname )
                elif Format == "PE" :
                    PEList.append( fname )
                else :
                    UnknownList.append( fname )
            
            FileList["PDF"] = PDFList
            FileList["OLE"] = OLEList
            FileList["PE"] = PEList
            FileList["Unknown"] = UnknownList
            
            return True
            
        except :
            print traceback.format_exc()
            return False
        
    
    def Separation(self, Options, FormatList, Format, ExceptList):
        try : 
            if Options.dst :
                dstdir = Options.dst 
            elif Options.directory :
                dstdir = Options.directory 
            else :
                print "Do not Sample's Directory"
                return []
            
            # Detailed Separation
            if Format == "OLE" : 
                OfficeList = []
                HWPList = []
                
                for fname in FormatList :
                    File = {} 
                    File["fname"] = fname
                    File["pBuf"] = FileControl.ReadFileByBinary(fname)
                    File['logbuf'] = ""
                    
                    OLE = OLEStruct( File )
                    
                    if not OLE.OLEHeader(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLEHeader()" )
                        continue
    
                    if not OLE.OLETableSAT(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLETableSAT()" )
                        continue
    
                    if not OLE.OLETableSSAT(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLETableSSAT()" )
                        continue
            
                    if not OLE.OLEDirectory(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLEDirectory()" )
                        continue
                    
                    if File["format"] == "Office" :
                        OfficeList.append( fname )
                    elif File["format"] == "HWP" :
                        HWPList.append( fname )
                    else :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : None Format" )
                        continue
                                
                if not self.SeparateFile(dstdir, HWPList, "HWP") :
                    print fname + "\tFailure : SeparateFile( HWP )"
                
                if not self.SeparateFile(dstdir, OfficeList, "Office") :
                    print fname + "\tFailure : SeparateFile( Office )"
                    
            else :
                if not self.SeparateFile(dstdir, FormatList, Format) :
                    print fname + "\tFailure : SeparateFile( %s )" % Format
        
        except : 
            print fname + "\tException : Separation()"
            print traceback.format_exc()


    def SeparateFile(self, curdirpath, flist, Format):
        try :
            dirpath = curdirpath + "\\" + Format
            if not os.path.exists( dirpath ) : 
                os.mkdir( dirpath )    
            
            for fname in flist :
                tmpcurdirpath = curdirpath + "\\" + fname
                tmpdirpath = dirpath + "\\" + fname
                shutil.move(tmpcurdirpath, tmpdirpath)
            
        except :
            print traceback.format_exc()
            return False
    
        return True


    def VulScan(self, OptScan, dstdir, FormatList):
        try :
            ScanList = {}
            dpath = []
            
            os.chdir( dstdir ) 
            
            for relativepath in os.listdir( dstdir ) :
                dpath.append( os.path.abspath(relativepath) )
            
            for dirpath in dpath :
                if os.path.isdir( dirpath ) and ( os.path.split(dirpath)[1] in FormatList ) :
                    if (OptScan != "*") and (os.path.split(dirpath)[1] == OptScan) :
                        ScanList[os.path.split(dirpath)[1]] = os.listdir( dirpath )
                    
                    if OptScan == "*" :
                        ScanList[os.path.split(dirpath)[1]] = os.listdir( dirpath )
            
            main = Main()
            
            for Format in FormatList :
                print Format
                print ScanList[Format]
            
            if not main.Scan(ScanList, FormatList) :            
                return False
            
            return True
        
        except :
            print traceback.format_exc()


class Main():
    def CheckFormat(self, fname):
        try :
            Format = ""
            
            pBuf = FileControl.ReadFileByBinary(fname)
            if pBuf == "" :
                return Format
            
            for FormatFunc in ScanFormatFunc :
                Format = FormatFunc.Check( pBuf )
                if Format != "" :
                    break
            
            return Format
            
        except :
            print traceback.format_exc()
            return Format


    def Scan(self, ScanList, FormatList):
        try :
            for Format in FormatList :
                if ScanList[Format] == [] :
                    continue
                
                os.chdir( Format )
                
                for fname in ScanList[Format] :
                    File = {}
                    File["fname"] = fname
                    File["pBuf"] = FileControl.ReadFileByBinary(fname)
                    File["logbuf"] = ""
                    
                    if Format == "PDF" :
                        File["format"] = "PDF"
                        PDFScan.Scan(File)
                        continue
                    
                    if Format == "OLE" :
                        File["format"] = "OLE"
                        OLEScan.Scan(File)
                        continue
            
            return True
        
        except :
            print traceback.format_exc()
            return False


ScanFormatFunc = [PDFScan, OLEScan, PEScan]


if __name__ == '__main__' :
    
    Init = Initialize()
    Options = Init.GetOptions()
    
    # Flow log
    log = ""
    
    try :
        print "[*] Start"
        
        Act = Action()
        
        # Options "Directory"
        if Options.directory :
            if not os.path.isdir( Options.directory ) :
                print "    [OptErr] Check please Directory Options ( %s )" % Options.directory
                exit(-1)
                
            os.chdir( Options.directory )
            
        
        # Options "DELETE"
        if Options.directory and Options.delete :
            if not Act.OptDelete(Options) :
                print 
                exit(-1)
    
        sFlag = False
    
        # Options "FTP"
        if Options.ip :
            if not Act.OptFTP(Options) :
                exit(-1)
                
            sFlag = True
            FileList = Act.Categorizer(Options)
            if FileList == {} :
                exit(-1)

        
        FormatList = ["PDF", "OLE"]
        # Options "Scan"
        if Options.scan :
            if Options.dst :
                dstdir = Options.dst
            elif Options.directory :
                dstdir = Options.directory
            else :
                exit(-1)
            
            if sFlag == False :
                FileList = Act.Categorizer(Options)
                if FileList == {} :
                    exit(-1)
                        
            if not Act.VulScan(Options.scan, dstdir, FormatList) :
                exit(-1)
            
    except :
        print traceback.format_exc()
    
    finally: 
        # Reault Log
        log = ""
        log = "=" * 70 + "\n" \
            + "     Result\n" \
            + "-" * 70 + "\n" \
            + "  PDF Files       : %d\n" % len(FileList["PDF"]) \
            + "  OLE Files       : %d\n" % len(FileList["OLE"]) \
            + "  PE Files        : %d\n" % len(FileList["PE"]) \
            + "  Unknown Files   : %d\n" % len(FileList["Unknown"]) \
            + "\n  File Count      : %d\n" % (len(FileList["PDF"]) + len(FileList["OLE"]) + len(FileList["PE"]) + len(FileList["Unknown"])) \
            + "-" * 70 + "\n"
            
        ExceptList = FileList["Except"]
        if len(ExceptList) :
            log += "  Except Files : %d\n" % (len(ExceptList)/2) \
                + "\n  [ File Name ]\t\t\t\t[ Description ]\n"
                
            index = 0
            while index < len(ExceptList) :
                log += "  %s\t%s" % (ExceptList[index], ExceptList[index+1])
                index += 2
    
    exit(0)
        
    
    
    
    
    
    
    
    
    
    
    
    
    