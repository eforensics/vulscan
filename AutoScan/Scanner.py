# -*- coding:utf-8 -*-

import os, sys
from ftplib import FTP
from shutil import move
from optparse import OptionParser
from traceback import format_exc


from Crawler import CFTPServer
from Common import CFile



from PDF import CPDF
from OLE import COLE
#from PE import CPE
#from HWP import CHWP23
#from RTF import CRTF


class CInitialize():
    #    Example)
    #        Crawling  : Scanner.py --ip <IP> --id <ID> --pw <PW> --src <SourceDirectory> --dst <DestinationDirectory>
    #        Delete    : Scanner.py -d <DestinationDirectory> --delete <FileExtend>
    #        Scan      : Scanner.py -d <DestinationDirectory> --scan <FileFormat>
    def fnGetOptions(self):
        Parser = OptionParser(usage='Crawling - %prog [--ip] IP [--id] LogID [--pw] LogPW [--src] SrcDir [--dst] DstDir\n\
                                     Delete   - %prog [-d,--directory] DstDir [--delete] FileExtend\n\
                                     DirScan  - %prog [-d,--directory] DstDir [--scan] FileFormat')
        # Crawling
        Parser.add_option('--ip', help='< IP >')
        Parser.add_option('--id', help='< ID >')
        Parser.add_option('--pw', help='< Password >')
        Parser.add_option('--src', help='< Absolutely Source Directory in Server >')
        Parser.add_option('--dst', help='< Absolutely Destination Directory in Client >')
        
        # Delete & Scan
        Parser.add_option('-d', '--directory', help='< Standard Directory >')
        Parser.add_option('--delete', help='< File Extend >')
        Parser.add_option('--scan', help='< Format >')
        Parser.add_option('--log', help='< Print File Struct >')
        (options, args) = Parser.parse_args()
            
        if len( sys.argv ) < 2 :
            Parser.print_help()
            exit(0)
    
        return options


class CAction():
    #    s_Extend        [IN]
    #    l_flist         [IN]
    #    Type BOOL       [OUT]        True / False
    def fnOptDelete(self, s_Extend, l_flist):
        # Opt
        #    Case 1. test.txt ( 'test', '.txt' )
        #    Case 2. *        ( '*', '' )
        #    Case 3. *.extend ( '*', '.extend' )
        print "[+] OptDelete........",
        
        try :
            l_OptExt = os.path.splitext( s_Extend )
            
            # Case 1
            if l_OptExt[0] != '*' :
                for s_fname in l_flist :
                    os.remove( s_fname )
                print "Success"
                return True
            
            # Case 2
            if l_OptExt[0] == '*' and l_OptExt[1] == '' :
                for s_fname in l_flist :
                    if os.path.isdir( s_fname ) :
                        os.removedirs( s_fname )
                        continue
                     
                    if os.path.isfile( s_fname ) :
                        os.remove( s_fname )
                        continue
                    
                print "Success"
                return True
            
            # Case 3
            if l_OptExt[1] != '' :
                for s_fname in l_flist :
                    if os.path.isdir( s_fname ) :
                        continue
                    
                    l_fext = os.path.splitext( s_fname )
                    if l_fext[1] == l_OptExt[1] :
                        os.remove( s_fname )
                        
                print "Success"
                return True
            
        except :
            print "\n\t[Failure] %s" % s_fname
            print format_exc()
            return False
            
            
    #    IP            [IN]        IP Address
    #    ID            [IN]        Login ID
    #    PW            [IN]        Login PW
    #    Src           [IN]        Source Directory      ( Server-Side )
    #    Dst           [IN]        Destination Directory ( Client-Side )
    #    Type BOOL     [OUT]       True / False
    def fnOptFTP(self, IP, ID, PW, Src, Dst):
        
        print "[+] OptFTP......"
        
        try :
            # Check Parameter
            if not IP or not ID or not PW or not Src or not Dst :
                print "\n\t[Failure] IP : %s" % IP      \
                    + "\n\t          ID : %s" % ID      \
                    + "\n\t          PW : %s" % PW      \
                    + "\n\t          Src : %s" % Src    \
                    + "\n\t          Dst : %s" % Dst 
                return False
            
            # Create Object : "FTP" & "Server" 
            FTP = FTP()
            FTPServer = CFTPServer()
            
            # Connection FTP Server
            if not FTPServer.fnConnectServer(FTP, IP, ID, PW) :
                return False
            
            # Download Files From FTP Server
            if not FTPServer.fnDownload(FTP, Src, Dst) :
                return False
            
            # Connection Termination
            FTP.close()
        
            return True
            
        except :
            print format_exc()
            return False


    #    None            [IN]
    #    BOOL Type       [OUT]        True/False
    def fnCategorize(self):
        
        print "[+] Classification "
        
        try :
            d_FileList = {}
            
            for s_Format in g_FileFormatList :
                d_FileList[ s_Format ] = []            
    
            # Check File Format
            if not self.fnCheckFormat( d_FileList ) :
                return False
            
            # Copy Files By File Format
            if not self.fnCopyFileByFormat( d_FileList ) :
                return False
            
            return True
            
        except :
            print format_exc()
            return False
    
    
    #    d_FileList        [IN]        File List in Directory
    #    BOOL Type         [OUT]       True/False
    def fnCheckFormat(self, d_FileList):
        
#        print "\t[-] CheckFormat"
        
        try :
            s_Format = ""
            l_curDirList = os.listdir( os.path.abspath(os.curdir) )
            
            for s_fname in l_curDirList :
                if os.path.isdir( s_fname ) :
                    continue
                
                s_pBuf = CFile.fnReadFile(s_fname)
                for s_Check in g_ClsCheck :
                    s_Format = eval("C"+s_Check).fnCheck(s_pBuf)
                    if s_Format in g_FileFormatList :
                        d_FileList[ s_Format ].append( s_fname )
                        break
                
                if s_Format == None :
                    print "\t\t[Except] Format Check ( %s )" % s_fname
                    d_FileList["Except"].append( s_fname )
                    
    
#            print "\t" + "=" * 40
#            for s_Format in g_FileFormatList :
#                print repr( s_Format ).rjust(15), repr( hex(len(d_FileList[ s_Format ])) ).rjust(20)
#            print "\t" + "=" * 40
                            
            return True
        
        except :
            print format_exc()
            return False
    
    
    #    d_FileList            [IN]        File List in Directory
    #    BOOL Type             [OUT]       True/False
    def fnCopyFileByFormat(self, d_FileList):
        
#        print "\t[-] CopyFile By Format"
        
        try :
            for s_Format in g_FileFormatList :
                s_DstDir = ".\\" + s_Format
                if not os.path.exists( s_DstDir ) :
                    os.mkdir( s_DstDir )
                
                for s_fname in d_FileList[ s_Format ] :
                    move( s_fname, ".\\" + s_Format )
            
            return True
    
        except :
            print format_exc()
            return False    
    
    
    #    OptScan                [IN]        Scan Option ( *, Format String )
    #    BOOL Type              [OUT]       True/False
    def fnOptScan(self, OptScan):
       
        print "[+] Scanning By File Format"
        
        try :
            l_EnableList = g_FileFormatList
            if OptScan != "*" : 
                l_EnableList = []
                l_EnableList.append( OptScan )            
            
            s_UpperPath = os.path.abspath( os.curdir )
            
            for s_Format in l_EnableList :
                if s_Format in g_ExceptScan :
                    continue
                
                if not s_Format in g_ClsCheck :
                    continue
                
                s_SubPath = s_UpperPath + "\\" + s_Format
                os.chdir( s_SubPath )
                
                l_SubPathFileList = os.listdir( s_SubPath )
                for s_fname in l_SubPathFileList :
                    if os.path.isdir( s_fname ) :
                        continue
                    
                    eval("C" + s_Format).fnScan( s_fname, CFile.fnReadFile(s_fname), s_Format )
                        
            return True
        
        except :
            print format_exc()
            return False



# Check File Object
g_ClsCheck = ["PDF"]
#g_ClsCheck = ["OLE", "PDF", "PE", "HWP23", "RTF"]

# Enable File Object 
g_FileFormatList = ["OLE", "PDF", "HWP23", "RTF", "PE", "Unknown", "Except"]

# Excepted Scan File Object
g_ExceptScan = ["OLE", "HWP23", "RTF", "PE", "Unknown", "Except"]
#------------------------------------------------------------------
# PDF
# 
# OLE
#  ├ HWP
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
# RTF
# 
# Unknown
# Except
#------------------------------------------------------------------

if __name__ == '__main__' :
    
    Initialize = CInitialize()
    Options = Initialize.fnGetOptions()
    
    print "[+] Start"
    
    try :
        Action = CAction()
        
        # Default Set Current Directory
        # - Delete or Scan
        if Options.directory :
            if not os.path.isdir( Options.directory ) :
                exit(-1)
            s_curdir = Options.directory
        # - Crawling
        elif Options.dst :
            if not os.path.isdir( Options.dst ) :
                exit(-1)
            s_curdir = Options.dst 
        else :
            exit(-1)
        
        os.chdir( s_curdir )
        
        # Options "DELETE"
        if Options.delete :
            l_flist = os.listdir( os.curdir )
            if not Action.fnOptDelete( Options.delete, l_flist ) :
                exit(-1)
        
        
        # Options "FTP"
        if Options.ip :
            if not Action.fnOptFTP( Options.ip, Options.id, Options.pw, Options.src, Options.dst ) :
                exit(-1)

        
        # Options "SCAN"
        if Options.scan :
            if not Action.fnCategorize() :
                exit(-1)
            
            if not Action.fnOptScan( Options.scan ) :
                exit(-1)
        
    except :
        print format_exc()
        exit(-1)
    
    exit(0)
    
    
    
    
    