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
        Parser.add_option('--src', help='< Source Directory in Server >')
        Parser.add_option('--dst', help='< Destination Directory in Client >')
        
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
    def ConnectServer(self, FTP, IP, ID, PW, log, Errlog):
        try :
            # Connect Network
            log += "[+] Connect Server : %s ( ID : %s ).........." % (IP, ID)
            MsgConnect = FTP.connect(IP)
            
            # Failure Connection
            #    - Success Message : 220 (vsFTPd 2.3.5)
            if MsgConnect.find("220") == -1 :
                Errlog += "Failure ( Connect : %s )\n" % MsgConnect
                return False
            
            # Failure Connection
            #    - Success Message : 230 Login successful
            MsgConnect = FTP.login(ID, PW)
            if MsgConnect.find("230") == -1 :
                Errlog += "Failure ( Login : %s )\n" % MsgConnect
                return False
            
            # Success Connection
            log += "Success\n"
            
        except :
            Errlog += traceback.format_exc()
            return False
        
        return True
    
    
    def Download(self, FTP, Src, Dst, log, Errlog):
        try :
            log += "[+] Download......"
            
            if not os.path.isdir( Dst ) :
                os.mkdir( Dst )
            
            os.chdir( Dst )
            SrcList = Src.split("\\")
            
            for nextdir in SrcList :
                MsgFTP = FTP.cwd(nextdir)
                if MsgFTP.find("250") == -1 :
                    Errlog += "Failure Change Directory\n"
                    return False            
            
            if not self.DownloadFile(FTP, log, Errlog) :
                return False
            
        except :
            Errlog += traceback.format_exc()
            return False
        
        return True
    
    
    def DownloadFile(self, FTP, log, Errlog):
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
                Errlog += "Folder is not Found\n"
                return True
            

            SrcPath = FTP.pwd()
            for extDir in ExtDir :
                # Move Src Directory ( in FTP Server )
                srcpath = SrcPath + "/" + extDir
                MsgFTP = FTP.cwd( srcpath )
                if MsgFTP.find("250") == -1 :
                    Errlog += "Failure Change Directory ( %s )" % extDir
                
                # File Download ( From FTP Server To Host PC )
                flist = FTP.nlst()
                for fname in flist :
                    if os.path.splitext( fname )[1] == ".txt" :
                        continue
                    
                    MsgFTP = FTP.retrlines("RETR " + fname, open(fname, 'wb').write)
                    if MsgFTP.find( "226" ) == -1 :
                        Errlog += "    %s ( %s )" % ( fname, MsgFTP )
                        continue
            
        except :
            Errlog += traceback.format_exc()
            return False
        
        return True
    
    
    def CloseServer(self, FTP, log, Errlog):
        try :
            FTP.close()
            log += "[+] Terminated FTP Connection\n"
        except :
            Errlog += traceback.format_exc()
            return False
        
        return True


class Action():
    def OptDelete(self, Options, log, Errlog):
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
            Errlog += traceback.format_exc()
            return False

    def OptFTP(self, Options, log, Errlog) :
        try : 
            IP = Options.ip
            ID = Options.id
            PW = Options.pw
            srcdir = Options.src
            dstdir = Options.dst
            
            if not IP or not ID or not PW or not srcdir or not dstdir :
                Errlog += "\n[OptErr] Check please FTP Options"
                return False    
            
            # Connection FTP Server
            FTP = ftplib.FTP()
            Server = FTPServer()
            tmplog = ""
            if not Server.ConnectServer(FTP, IP, ID, PW, log, tmplog) :
                Errlog += tmplog
                return False
            
            log += tmplog
            
            # Download Files From FTP Server
            tmplog = ""
            if not Server.Download(FTP, srcdir, dstdir, log, tmplog) :
                Errlog += tmplog
                return False
            
            log += tmplog
            
            # Connection Termination
            tmplog = ""
            Server.CloseServer(FTP, tmplog)
            log += tmplog
        
            return True
        except :
            Errlog += traceback.format_exc()
            return False


    def isCategorizer(self, dstdir, FormatList, Errlog):
        try :
            predir = os.path.split( dstdir )
            if predir[1] in FormatList :
                return True
            else :
                return False
        
        except :
            Errlog += traceback.format_exc()
            return False


    def Categorizer(self, Options, log, Errlog):
        try :    
            # Separate File Format
            FileList = {}
            log += "[+] Check Samples File Format........"
            Errlog = "Into SeparateList()\n"
            FileList = self.SeparateList(Options, FileList, log, Errlog)
            if FileList == {} :
                return {}
                
            # Separate Samples
            log += "[+] Separate Files.........\n"
            
            Errlog = "Into Separation()\n"
            if FileList["PDFList"] != [] :
                self.Separation(Options, FileList["PDF"], "PDF", log, Errlog) 
                
            if FileList["OLEList"] != [] :
                self.Separation(Options, FileList["OLE"], "OLE", log, Errlog) 
                
            if FileList["PEList"] != [] :
                self.Separation(Options, FileList["PE"], "PE", log, Errlog)
            
            if FileList["UnknownList"] != [] :
                self.Separation(Options, FileList["Unknown"], "unknown", log, Errlog) 

            return FileList
        
        except :
            Errlog += traceback.format_exc()
            return {}


    def SeparateList(self, Options, FileList, log, Errlog):
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
                Errlog += "Do not Sample's Directory"
                return []
            
            main = Main()
            os.chdir( dstdir )
            flist = os.listdir( os.curdir )
            for fname in flist :
                if os.path.isdir( fname ) :
                    continue
                
                Format = main.CheckFormat(fname, log, Errlog)
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
            
            return FileList
            
        except :
            Errlog += traceback.format_exc()
            return []
        
    
    def Separation(self, Options, FormatList, Format, log, Errlog):
        try : 
            if Options.dst :
                dstdir = Options.dst 
            elif Options.directory :
                dstdir = Options.directory 
            else :
                Errlog += "Do not Sample's Directory"
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
                        Errlog += fname + "\tFailure : OLEHeader()\n"
                        continue
    
                    if not OLE.OLETableSAT(File) :
                        Errlog += fname + "\tFailure : OLETableSAT()\n"
                        continue
    
                    if not OLE.OLETableSSAT(File) :
                        Errlog += fname + "\tFailure : OLETableSSAT()\n"
                        continue
            
                    if not OLE.OLEDirectory(File) :
                        Errlog += fname + "\tFailure : OLEDirectory()\n"
                        continue
                    
                    if File["format"] == "Office" :
                        OfficeList.append( fname )
                    elif File["format"] == "HWP" :
                        HWPList.append( fname )
                    else :
                        Errlog += fname + "\tFailure : None Format\n"
                        continue
                
                if not self.SeparateFile(dstdir, HWPList, "HWP") :
                    Errlog += fname + "\tFailure : SeparateFile( HWP )"
                
                if not self.SeparateFile(dstdir, OfficeList, "Office") :
                    Errlog += fname + "\tFailure : SeparateFile( Office )"
                    
            else :
                if not self.SeparateFile(dstdir, FormatList, Format) :
                    Errlog += fname + "\tFailure : SeparateFile( %s )" % Format
        
        except : 
            Errlog += fname + "\tException : Separation()"
            Errlog += traceback.format_exc()


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


    def VulScan(self, OptScan, FormatList, log, Errlog):
        try :
            ScanList = {}
            dpath = []
            
            for relativepath in os.listdir( os.curdir ) :
                dpath.append( os.path.abspath(relativepath) )
                 
            for dirpath in dpath :
                if os.path.isdir( dirpath ) and ( os.path.split(dirpath)[1] in FormatList ) :
                    if (OptScan != "*") and (os.path.split(dirpath)[1] == OptScan) :
                        ScanList[os.path.split(dirpath)[1]] = os.listdir( dirpath )
                        break
                    
                    if OptScan == "*" :
                        ScanList[os.path.split(dirpath)[1]] = os.listdir( dirpath )
            
            main = Main()
            if not main.Scan(ScanList, FormatList, log, Errlog) :            
                return False
            
            return True
        
        except :
            Errlog += traceback.format_exc()


class Main():
    def CheckFormat(self, fname, log, Errlog):
        try :
            File = {}
            
            pBuf = FileControl.ReadFileByBinary(fname, Errlog)
            if pBuf == "" :
                return False
            
            File["fname"] = fname
            File["pBuf"] = pBuf
            
            Format = ""
            for FormatFunc in ScanFormatFunc :
                Format = FormatFunc.Check()
                if Format != "" :
                    break
            
        except :
            Errlog += traceback.format_exc()
            return False


    def Scan(self, ScanList, FormatList, log, Errlog):
        try :
            for Format in FormatList :
                if ScanList[Format] == [] :
                    continue
                
                for fname in ScanList[Format] :
                    File = {}
                    File["fname"] = fname
                    File["pBuf"] = FileControl.ReadFileByBinary(fname, Errlog)
                    if Format == "PDF" :
                        PDFScan.Scan(File)
                    elif Format == "OLE" :
                        OLEScan.Scan(File)
                    else :
                        break
            
            return True
        
        except :
            Errlog += traceback.format_exc()
            return False


ScanFormatFunc = [PDFScan, OLEScan, PEScan]


if __name__ == '__main__' :
    
    Init = Initialize()
    Options = Init.GetOptions()
    
    # Flow log
    log = ""
    # Except/Error log in to Function
    Errlog = ""
    
    try :
        Act = Action()
        
        # Options "Directory"
        if Options.directory :
            if not os.path.isdir( Options.directory ) :
                Errlog += "[OptErr] Check please Directory Options ( %s )" % Options.directory
                exit(-1)
                
            os.chdir( Options.directory )
            
        
        # Options "DELETE"
        if Options.directory and Options.delete :
            Errlog = "Into Options \"Delete()\"\n"
            if not Act.OptDelete(Options, log, Errlog) :
                exit(-1)
    
    
        FormatList = ["PDF", "OLE"]
    
        # Options "FTP"
        if Options.ip :
            Errlog = "Into Options \"FTP()\"\n"
            if not Act.OptFTP(Options, log, Errlog) :
                exit(-1)
            
            FileList = Act.Categorizer(Options, log, Errlog)
            if FileList == {} :
                exit(-1)
        
        
        # Options "Scan"
        if Options.scan :
            Errlog = "Into Options \"Scan()\"\n"
            
            if Options.dst :
                dstdir = Options.dst
            elif Options.directory :
                dstdir = Options.directory
            else :
                exit(-1)
            
            if not Act.isCategorizer(dstdir, FormatList, Errlog) :
                FileList = Act.Categorizer(Options, log, Errlog)
                if FileList == {} :
                    exit(-1)
            
            if not Act.VulScan(Options.scan, FormatList, log, Errlog) :
                exit(-1)
        
        
        # Reault Log
        log += "=" * 70 + "\n" \
            + "     Result\n" \
            + "-" * 70 + "\n" \
            + "  PDF Files       : %d\n" % len(FileList["PDF"]) \
            + "  OLE Files       : %d\n" % len(FileList["OLE"]) \
            + "  PE Files        : %d\n" % len(FileList["PE"]) \
            + "  Unknown Files   : %d\n" % len(FileList["Unknown"]) \
            + "\n  File Count      : %d\n" % (len(FileList["PDF"]) + len(FileList["OLE"]) + len(FileList["PE"]) + len(FileList["Unknown"])) \
            + "-" * 70 + "\n"
        
#        if len(ExceptList) :
#            log += "  Except Files : %d\n" % (len(ExceptList)/2) \
#                + "\n  [ File Name ]\t\t\t\t[ Description ]\n"
#            
#            index = 0
#            while index < len(ExceptList) :
#                log += "  %s\t%s" % (ExceptList[index], ExceptList[index+1])
#                index += 2
                
    
    except :
        Errlog += "\n" + traceback.format_exc()
        print Errlog
    
    finally :
        print log
        
    
    exit(0)
        
        
    
    
    
    
    
    
    
    
    
    
    
    
    