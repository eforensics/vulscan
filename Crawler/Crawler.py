# -*- coding:utf-8 -*-

# import Public Module
import optparse, traceback, sys, os, re, struct, ftplib, types, shutil


# import Private Module
from ComFunc import FileControl, BufferControl
from OLEScanner import OLEScan
from PDFScanner import PDFScan

class Initialize():
    def GetOption(self):
        Parser = optparse.OptionParser(usage='usage: %prog [--IP] IP [--ID] LogID [--PW] LogPW [--src] SrcDir [--dst] DstDir\n')
        Parser.add_option('--ip', help='< IP >')
        Parser.add_option('--id', help='< ID >')
        Parser.add_option('--pw', help='< Password >')
        Parser.add_option('--src', help='< Source Directory >')
        Parser.add_option('--dst', help='< Destination Directory >')
        
        Parser.add_option('--dir','--directory', help='< delete directory >')
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
            pBuf = FileControl.ReadFileByBinary(fname)
            if PDFScan.Check(pBuf) :
                return "PDF"
            elif OLEScan.Check(pBuf) :
                return "OLE"
            else :
                return ""
            
        except :
            print traceback.format_exc()   


    def SeparateFile(self, curdirpath, flist, Format):
        try :
            dirpath = curdirpath + "\\" + Format
            
            if not os.path.exists( dirpath ) : 
                os.mkdir( dirpath )    
            
            curdirpath += "\\" + fname
            dirpath += "\\" + fname
            shutil.move(curdirpath, dirpath)
            
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
            TargetName = ["PDF", "unknown", "MS Compress", "MS Excel Spreadsheet", "MS Word Document", "Office Open XML Document"]
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
        # Option : Delete Files 
        if Options.delete and Options.directory :
            flist = os.listdir( Options.directory )
            for fname in flist :
                fext = os.path.splitext( fname )
                fdel = os.path.splitext( Options.delete )
                
                if fext[1] == fdel[1] :
                    os.remove( fname )
        
            exit(0)
        
        
        IP = Options.ip
        ID = Options.id
        PW = Options.pw
        SrcDir = Options.src
        DstDir = Options.dst
        
        log = ""
        
        
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
        HWPList = []
        OfficeList = []
        NoneSupport = []
        
        main = Main()
        flist = os.listdir( DstDir )
        for fname in flist : 
            Format = main.CheckFormat(fname)
            if Format == "PDF" :
                PDFList.append( fname )
            elif Format == "Office" :
                OfficeList.append( fname )
            elif Format == "HWP" :
                HWPList.append( fname )
            else :
                NoneSupport.append( fname )
        
        if PDFList == [] and HWPList == [] and OfficeList == [] :
            log += "None\n"
        else :
            log += "Done\n"
        

        # Separate Samples
        log += "[+] Separate Files.........\n"
        if PDFList != [] :
            if not main.SeparateFile(DstDir, PDFList, "PDF") :
                log += "    Failure Separate PDF\n"
            
        if HWPList != [] :
            if not main.SeparateFile(DstDir, HWPList, "HWP") :
                log += "    Failure Separate HWP\n"
            
        if OfficeList != [] :
            if not main.SeparateFile(DstDir, OfficeList, "Office") :
                log += "    Failure Separate Office\n"
        
        if NoneSupport != [] :
            if not main.SeparateFile(DstDir, NoneSupport, "unknown") :
                log += "    Failure Separate Unknown\n"
        
        # Reault Log
        log += "=======================================\n" \
            + " Result\n" \
            + "---------------------------------------\n" \
            + "PDF Files    : %d\n" % len(PDFList) \
            + "HWP Files    : %d\n" % len(HWPList) \
            + "Office Files : %d\n" % len(OfficeList) \
            + "None Files   : %d\n" % len(NoneSupport) \
            + "Total Count  : %d\n" % len(flist) \
            + "---------------------------------------\n"
#            + "\nMove Count : %d\n" % (len(PDFList) + len(HWPList) + len(OfficeList)) \
    
    except :
        print traceback.format_exc() 
    
    finally:
        if isinstance( FTP, types.InstanceType ) :
            Server.CloseServer(FTP, log)
            
        print log
    
    exit(0)












