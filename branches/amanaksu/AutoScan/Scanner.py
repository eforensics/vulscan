# -*- coding:utf-8 -*-

# import Public Module
import optparse, traceback, sys, os, ftplib, shutil

# import Private Module
import Common
from PDFScanner import PDF
from OLEScanner import OLE
from PEScanner import PE
from HWPScanner import HWP23




class Initialize():
    def GetOptions(self):
        Parser = optparse.OptionParser(usage='usage1: %prog [--IP] IP [--ID] LogID [--PW] LogPW [--src] SrcDir [--dst] DstDir\n\nusage2: %prog [-d,--directory] DstDir [--scan] * or FileFormat')
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
            print "\t[-] Connect Server : %s.........." % IP,
            MsgConnect = FTP.connect(IP)
            
            # Failure Connection
            #    - Success Message : 220 (vsFTPd 2.3.5)
            if MsgConnect.find("220") == -1 :
                print "\t\t[Failure] %s" % MsgConnect
                return False
            
            # Failure Connection
            #    - Success Message : 230 Login successful
            MsgConnect = FTP.login(ID, PW)
            if MsgConnect.find("230") == -1 :
                print "\t\t[Failure] %s" % MsgConnect
                return False
            
            print "Success"
            return True
            
        except :
            print traceback.format_exc()
            return False
    
    
    def Download(self, FTP, Src, Dst):
        try :
            print "\t[-] Download"
            
            # Set Destination Folder 
            if not os.path.isdir( Src ) :
                print "\t\t[Failure] Don't Exist Folder ( %s )" % Src
                return False
            
            if not os.path.isdir( Dst ) :
                os.mkdir( Dst )
            
            os.chdir( Dst )
            
            # Set Source Folder
            # Ex) 2012/10/5 -> '2012', '10', '5'
            SrcList = Src.split("\\")
            
            for nextdir in SrcList :
                MsgFTP = FTP.cwd(nextdir)
                if MsgFTP.find("250") == -1 :
                    print "\t\t[Failure] Change Directory ( %s )" % nextdir
                    return False            
            
            # File Download
            if not self.DownloadFile(FTP) :
                return False
            
            return True
            
        except :
            print traceback.format_exc()
            return False
        
    
    def DownloadFile(self, FTP):
        
        print "\t\t[-] DownloadFile"
        
        try :
            # Check Target Directory By File Format
            TargetName = ["PDF",        \
                          "unknown",    \
                          "MS Compress", "MS Excel Spreadsheet", "MS Word Document", "Office Open XML Document", "Office Open XML Spreadsheet"]
            ExtDir = []
            
            DirList = FTP.nlst()
            for DirName in DirList : 
                if DirName in TargetName :
                    ExtDir.append( DirName )
            
            if ExtDir == [] :
                print "\t\t\t[Failure] Target Folder is not Found"
                return False

            CopyFlag = False
            SrcPath = FTP.pwd()
            for extDir in ExtDir :
                # Move Src Directory ( in FTP Server )
                srcpath = SrcPath + "/" + extDir
                MsgFTP = FTP.cwd( srcpath )
                if MsgFTP.find("250") == -1 :
                    print "\t\t\t[Failure] Change Directory ( %s )" % extDir
                    continue
                
                # File Download ( From FTP Server To Host PC )
                FileCount = 0
                flist = FTP.nlst()
                for fname in flist :
                    if os.path.splitext( fname )[1] == ".txt" :
                        continue
                    
                    MsgFTP = FTP.retrlines("RETR " + fname, open(fname, 'wb').write)
                    if MsgFTP.find( "226" ) == -1 :
                        print "\t\t\t[Failure] RETRLines %s ( %s )" % ( fname, MsgFTP )
                        continue
                    
                    FileCount += 1
                
                if FileCount > 0 :
                    print "\t\t\t%s ( Count : %d )" %  ( extDir, FileCount)
                    CopyFlag = True
            
            if CopyFlag :
                return True
            else :
                return False
                        
        except :
            print traceback.format_exc()
            return False
        

class Action():
    # Opt
    #    Case 1. test.txt ( 'test', '.txt' )
    #    Case 2. *        ( '*', '' )
    #    Case 3. *.extend ( '*', '.extend' )
    def OptDelete(self, Opt, flist):
        
        print "[+] OptDelete........",
        
        try :
            OptExt = os.path.splitext( Opt )
            
            # Case 1
            if OptExt[0] != "*" :
                for fname in flist :
                    os.remove( fname )
                
                print "Success"
                return True
            
            # Case 2
            if OptExt[0] == "*" and OptExt[1] == "" :
                for fname in flist :
                    if os.path.isdir( fname ) :
                        os.removedirs( fname )
                        continue
                     
                    if os.path.isfile( fname ) :
                        os.remove( fname )
                        continue
                
                print "Success"
                return True
                
            # Case 3
            if OptExt[1] != "" :        
                for fname in flist :
                    if os.path.isdir( fname ) :
                        continue
                    
                    fext = os.path.splitext( fname )
                    if fext[1] == OptExt[1] :
                        os.remove( fname )
            
                print "Success"
                return True
        
        except :
            print "\n\t[Failure] %s" % fname 
            print traceback.format_exc()
            return False       
    
    
    def OptFTP(self, IP, ID, PW, Src, Dst):
        
        print "[+] OptFTP......"
        
        try :
            # Check Parameter
            if not IP or not ID or not PW or not Src or not Dst :
                print "\n\t[Failure] IP : %s" % IP          \
                        + "\n\t          ID : %s" % ID      \
                        + "\n\t          PW : %s" % PW      \
                        + "\n\t          Src : %s" % Src    \
                        + "\n\t          Dst : %s" % Dst 
                return False
            
            # Create Object : "FTP" & "Server" 
            FTP = ftplib.FTP()
            Server = FTPServer()
            
            # Connection FTP Server
            if not Server.ConnectServer(FTP, IP, ID, PW) :
                return False
            
            # Download Files From FTP Server
            if not Server.Download(FTP, Src, Dst) :
                return False
            
            # Connection Termination
            FTP.close()
        
            return True
            
        except :
            print traceback.format_exc()
            return False
    

    def Categorizer(self):
        
        print "[+] Classification "
        
        try :
            # Check File Format
            FileList = {}
            
            for Format in FileFormatList :
                FileList[ Format ] = []            
    
            if not self.CheckFormat( FileList ) :
                return False
            
            
            # Copy Files By File Format
            if not self.CopyFileByFormat( FileList ) :
                return False
            
            return True
            
        except :
            print traceback.format_exc()
            return False
    
    
    def CheckFormat(self, FileList):
        
        print "\t[-] CheckFormat"
        
        try :
            Format = ""
            curDirList = os.listdir( os.path.abspath(os.curdir) )
            
            for fname in curDirList :
                if os.path.isdir( fname ) :
                    continue
                
                pBuf = Common.FileControl.ReadFileByBinary(fname)
                for mCheck in ClsCheck :
                    Format = eval(mCheck).Check(pBuf)
                    if Format in FileFormatList :
                        FileList[ Format ].append( fname )
                        break
                
                if Format == "" :
                    print "\t\t[Except] Format Check ( %s )" % fname
                    FileList["Except"].append( fname )
                    
    
            print "\t" + "=" * 40
            for Format in FileFormatList :
                print "\t   %s       \t: %d" % ( Format, len(FileList[ Format ]) )
            print "\t" + "=" * 40
                            
            return True
        
        except :
            print traceback.format_exc()
            return False
    
    
    def CopyFileByFormat(self, FileList):
        
        print "\t[-] CopyFile By Format"
        
        try :
            for Format in FileFormatList :
                DstDir = ".\\" + Format
                if not os.path.exists( DstDir ) :
                    os.mkdir( DstDir )
                
                for fname in FileList[ Format ] :
                    shutil.move( fname, ".\\" + Format )
            
            return True
    
        except :
            print traceback.format_exc()
            return False
    
    
    def OptScan(self, ScanOpt):
        
        print "[+] Scanning By File Format"
        
        try :
            EnableList = FileFormatList
            if ScanOpt != "*" : 
                EnableList = []
                EnableList.append( ScanOpt )            
            
            UpperPath = os.path.abspath( os.curdir )
            
            for Format in FileFormatList :
                if Format in ExceptScan :
                    continue
                
                if not Format in EnableList :
                    continue
                
                SubPath = UpperPath + "\\" + Format
                
                SubPathList = os.listdir( SubPath )
                for fname in SubPathList :
                    if os.path.isdir( fname ) :
                        continue
                    
                    os.chdir( SubPath )
                    
                    File = {}
                    File["fname"] = fname
                    File["pBuf"] = Common.FileControl.ReadFileByBinary(fname)
                    File["format"] = Format
                    
                    eval(Format).Scan( File )
                        
            return True
        
        except :
            print traceback.format_exc()
            return False


# Check File Object
ClsCheck = ["PDF", "OLE", "PE", "HWP23"]

# Enable File Object 
FileFormatList = ["PDF", "OLE", "HWP23", "PE", "Unknown", "Except"]

# Excepted Scan File Object
ExceptScan = ["HWP23", "PE", "Unknown", "Except"]
#------------------------------------------------------------------
# PDF
# 
# OLE
#  ├ HWP ( OLE )
#  └ Office
#      ├ DOC
#      ├ PowerPoint
#      └ Excel
# 
# HWP23 ( HWP2.0 or 3.0 )
#
# COFF
#  └ PE
# 
# Unknown
# Except
#------------------------------------------------------------------





if __name__ == '__main__' :
    
    Init = Initialize()
    Options = Init.GetOptions()
    
    print "[+] Start"
    
    try :
        Act = Action()
        
        # Default Set Current Directory
        if Options.directory :
            if not os.path.isdir( Options.directory ) :
                exit(-1)
            curdir = Options.directory
        elif Options.dst :
            if not os.path.isdir( Options.dst ) :
                exit(-1)
            curdir = Options.dst 
        else :
            exit(-1)
        
        os.chdir( curdir )
        
        
        # Options "DELETE"
        if Options.delete :
            flist = os.listdir( os.curdir )
            if not Act.OptDelete( Options.delete, flist ) :
                exit(-1)
        
        
        # Options "FTP"
        if Options.ip :
            if not Act.OptFTP( Options.ip, Options.id, Options.pw, Options.src, Options.dst ) :
                exit(-1)

        
        # Options "SCAN"
        if Options.scan :
            if not Act.Categorizer() :
                exit(-1)
            
            if not Act.OptScan( Options.scan ) :
                exit(-1)
        
    except :
        print traceback.format_exc()
        exit(-1)
    
    exit(0)
    
    
    
    
    
    
    
    
    
    