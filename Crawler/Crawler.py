# -*- coding:utf-8 -*-

# import Public Module
import optparse, traceback, sys, os, ftplib, shutil


# import Private Module
from ComFunc import FileControl
from OLEScanner import OLEScan, OLEStruct
from PDFScanner import PDFScan
from PEScanner import PEScan
from RTFScanner import RTFScan


class Initialize():
    def GetOption(self):
        Parser = optparse.OptionParser(usage='usage: %prog [--IP] IP [--ID] LogID [--PW] LogPW [--src] SrcDir [--dst] DstDir\n')
        Parser.add_option('--ip', help='< IP >')
        Parser.add_option('--id', help='< ID >')
        Parser.add_option('--pw', help='< Password >')
        Parser.add_option('--src', help='< Source Directory >')
        Parser.add_option('--dst', help='< Destination Directory >')
        
        Parser.add_option('--dir', help='< delete directory >')
        Parser.add_option('--delete', help='< Extend >')
        
        Parser.add_option('--log', help='< file name >')
        (options, args) = Parser.parse_args()
            
        if len( sys.argv ) < 4 :
            Parser.print_help()
            exit(0)
    
        return options


class Main():
    def CheckFormat(self, fname):    
        try :
            File = {}
            File["fname"] = fname
            File['logbuf'] = ""
            
            pBuf = FileControl.ReadFileByBinary(fname)
            File["pBuf"] = pBuf
            
            ClsList = [PDFScan, OLEScan, PEScan, RTFScan]
            
            Format = ""
            for Cls in ClsList :
                Format = Cls.Check(pBuf) 
                if Format != "" :
                    break
            
        except :
            print traceback.format_exc()
            
        return Format   


    def Separation(self, curdirpath, flist, Format, ExceptList, log):
        try : 
            # Detailed Separation
            if Format == "OLE" : 
                OfficeList = []
                HWPList = []
                
                for fname in flist :
                    File = {} 
                    File["fname"] = fname
                    File["pBuf"] = FileControl.ReadFileByBinary(fname)
                    File['logbuf'] = ""
                    
                    OLE = OLEStruct( File )
                    
                    if not OLE.OLEHeader(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLEHeader()\n")
                        continue
    
                    if not OLE.OLETableSAT(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLETableSAT()\n")
                        continue
    
                    if not OLE.OLETableSSAT(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLETableSSAT()\n")
                        continue
            
                    if not OLE.OLEDirectory(File) :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : OLEDirectory()\n")
                        continue
                    
                    if File["format"] == "Office" :
                        OfficeList.append( fname )
                    elif File["format"] == "HWP" :
                        HWPList.append( fname )
                    else :
                        ExceptList.append( fname )
                        ExceptList.append( "Failure : None Format\n")
                        continue
                
                if not self.SeparateFile(curdirpath, HWPList, "HWP") :
                    ExceptList.append( fname )
                    ExceptList.append( "Failure : SeparateFile( HWP )")
                    
                
                if not self.SeparateFile(curdirpath, OfficeList, "Office") :
                    ExceptList.append( fname )
                    ExceptList.append( "Failure : SeparateFile( Office )")
                
            else :
                if not self.SeparateFile(curdirpath, flist, Format) :
                    ExceptList.append( fname )
                    strlog = "Failure : SeparateFile( %s )" % Format
                    ExceptList.append( strlog )
        
        except : 
            ExceptList.append( fname )
            ExceptList.append( "Exception : Separation()" )
            print traceback.format_exc()
            return False
        
        return True


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


class FTPServer():
    def ConnectServer(self, FTP, IP, ID, PW, log):
        try :
            # Connect Network
            log += "[+] Connect Server : %s ( ID : %s ).........." % (IP, ID)
            MsgConnect = FTP.connect(IP)
            
            # Failure Connection
            #    - Success Message : 220 (vsFTPd 2.3.5)
            if MsgConnect.find("220") == -1 :
                log += "Failure ( Connect : %s )\n" % MsgConnect
                return False
            
            # Failure Connection
            #    - Success Message : 230 Login successful
            MsgConnect = FTP.login(ID, PW)
            if MsgConnect.find("230") == -1 :
                log += "Failure ( Login : %s )\n" % MsgConnect
                return False
            
            # Success Connection
            log += "Success\n"
            
        except :
            print traceback.format_exc()
            return False
        
        return True
    
    
    def Download(self, FTP, Src, Dst, log):
        try :
            log += "[+] Download......"
            
            if not os.path.isdir( Dst ) :
                os.mkdir( Dst )
            
            os.chdir( Dst )
            SrcList = Src.split("\\")
            
            for nextdir in SrcList :
                MsgFTP = FTP.cwd(nextdir)
                if MsgFTP.find("250") == -1 :
                    log += "Failure Change Directory\n"
                    return False            
            
            if not self.DownloadFile(FTP, log) :
                return False
            
        except :
            print traceback.format_exc()
            return False
        
        return True
    
    
    def DownloadFile(self, FTP, log):
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
                log += "Folder is not Found\n"
                return True
            

            SrcPath = FTP.pwd()
            for extDir in ExtDir :
                # Move Src Directory ( in FTP Server )
                srcpath = SrcPath + "/" + extDir
                MsgFTP = FTP.cwd( srcpath )
                if MsgFTP.find("250") == -1 :
                    log += "Failure Change Directory ( %s )" % extDir
                
                # File Download ( From FTP Server To Host PC )
                flist = FTP.nlst()
                for fname in flist :
                    if os.path.splitext( fname )[1] == ".txt" :
                        continue
                    
                    MsgFTP = FTP.retrlines("RETR " + fname, open(fname, 'wb').write)
                    if MsgFTP.find( "226" ) == -1 :
                        log += "    %s ( %s )" % ( fname, MsgFTP )
                        continue
            
        except :
            print traceback.format_exc()
            return False
        
        return True
    
    
    def CloseServer(self, FTP, log):
        try :
            FTP.close()
            log += "[+] Terminated FTP Connection\n"
        except :
            print traceback.format_exc()
            return False
        
        return True
    


if __name__ == '__main__' :
    
    Init = Initialize()
    Options = Init.GetOption()
    
    try :
        IP = Options.ip
        ID = Options.id
        PW = Options.pw
        SrcDir = Options.src
        DstDir = Options.dst
        
        log = ""
        
        # Option : Delete Files 
        if Options.delete and Options.dir :
            os.chdir( Options.dir )
            
            flist = os.listdir( Options.dir )
            for fname in flist :
                fext = os.path.splitext( fname )
                fdel = os.path.splitext( Options.delete )
                
                if fext[1] == fdel[1] :
                    os.remove( fname )
        
            exit(0)
        
        if not IP or not ID or not PW or not SrcDir or not DstDir :
            exit(-1)
        
        
        # Connection FTP Server
        FTP = ftplib.FTP()
        Server = FTPServer()
        if not Server.ConnectServer(FTP, IP, ID, PW, log) :
            exit(-1)
        
        # Download Files From FTP Server
        if not Server.Download(FTP, SrcDir, DstDir, log) :
            exit(-1)
        
        # Connection Termination
        Server.CloseServer(FTP, log)
        FTP = ""
        
        # Check Samples
        log += "[+] Check Samples File Format........"
        PDFList = []
        OLEList = []
        PEList = []
        RTFList = []
        NoneSupport = []
        
#        os.chdir( DstDir )
        
        DirCnt = 0
        main = Main()
        flist = os.listdir( DstDir )
        for fname in flist : 
            if os.path.isdir( fname ) :
                DirCnt += 1
                continue
            
            Format = main.CheckFormat(fname)
            if Format == "PDF" :
                PDFList.append( fname )
            elif Format == "OLE" :  # Office and HWP
                OLEList.append( fname )
            elif Format == "PE" :
                PEList.append( fname )
            elif Format == "RTF" :
                RTFList.append( fname )
            else :
                NoneSupport.append( fname )
        
        if PDFList == [] and OLEList == [] and RTFList == [] :
            log += "None\n"
        else :
            log += "Done\n"
        
        
        # Separate Samples
        log += "[+] Separate Files.........\n"
        ExceptList = []
        if PDFList != [] :
            main.Separation(DstDir, PDFList, "PDF", ExceptList, log) 
            
        if OLEList != [] :
            main.Separation(DstDir, OLEList, "OLE", ExceptList, log) 
            
        if PEList != [] :
            main.Separation(DstDir, PEList, "PE", ExceptList, log)
            
        if RTFList != [] :
            main.Separation(DstDir, PEList, "RTF", ExceptList, log)
        
        if NoneSupport != [] :
            main.Separation(DstDir, NoneSupport, "unknown", ExceptList, log) 
        
        # Reault Log
        log += "=" * 70 + "\n" \
            + "     Result\n" \
            + "-" * 70 + "\n" \
            + "  PDF Files       : %d\n" % len(PDFList) \
            + "  OLE Files       : %d\n" % len(OLEList) \
            + "  PE Files        : %d\n" % len(PEList) \
            + "  RTF Files       : %d\n" % len(RTFList) \
            + "  None Files      : %d\n" % len(NoneSupport) \
            + "\n  File Count      : %d\n" % (len(flist) - DirCnt) \
            + "  Directory Count : %d\n" % DirCnt \
            + "  Total Count     : %d\n" % len(flist) \
            + "-" * 70 + "\n"
        
        if len(ExceptList) :
            log += "  Except Files : %d\n" % (len(ExceptList)/2)
            
            index = 0
            while index < len(ExceptList) :
                log += "  %s\t%s" % (ExceptList[index], ExceptList[index+1])
                index += 2
    
    except :
        print traceback.format_exc() 
    
    finally:
        
        FileControl.WriteFile("%s\\%s.log" % (DstDir, os.path.split(DstDir)[1]), log)
    
    exit(0)

