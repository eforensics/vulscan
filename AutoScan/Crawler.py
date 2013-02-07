# -*- coding:utf-8 -*-

import os 

from traceback import format_exc



class CFTPServer():
    def fnConnectServer(self, FTP, IP, ID, PW):
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
            print format_exc()
            return False
    
    
    def fnDownload(self, FTP, Src, Dst):
        try :
            print "\t[-] Download"
            
            # Set Destination Folder 
            if not os.path.isdir( Dst ) :
                os.mkdir( Dst )
            
            os.chdir( Dst )
            
            # Set Source Folder
            # Ex) 2012/10/5 -> '2012', '10', '5'
            l_SrcList = Src.split("\\")
            
            for s_nextdir in l_SrcList :
                MsgFTP = FTP.cwd(s_nextdir)
                if MsgFTP.find("250") == -1 :
                    print "\t\t[Failure] Change Directory ( %s )" % s_nextdir
                    return False            
            
            # File Download
            if not self.fnDownloadFile(FTP) :
                return False
            
            return True
            
        except :
            print format_exc()
            return False
        
    
    def fnDownloadFile(self, FTP):
        
        print "\t\t[-] DownloadFile"
        
        try :
            # Check Target Directory By File Format
            l_TargetName = ["PDF",        \
                          "unknown",    \
                          "MS Compress", "MS Excel Spreadsheet", "MS Word Document", "Office Open XML Document", "Office Open XML Spreadsheet", \
                          "Rich Text Format"]
            l_ExtDir = []
            
            l_DirList = FTP.nlst()
            for s_DirName in l_DirList : 
                if s_DirName in l_TargetName :
                    l_ExtDir.append( s_DirName )
            
            if l_ExtDir == [] :
                print "\t\t\t[Failure] Target Folder is not Found"
                return False

            CopyFlag = False
            s_SrcPath = FTP.pwd()
            for s_extDir in l_ExtDir :
                # Move Src Directory ( in FTP Server )
                s_srcpath = s_SrcPath + "/" + s_extDir
                MsgFTP = FTP.cwd( s_srcpath )
                if MsgFTP.find("250") == -1 :
                    print "\t\t\t[Failure] Change Directory ( %s )" % s_extDir
                    continue
                
                # File Download ( From FTP Server To Host PC )
                FileCount = 0
                l_flist = FTP.nlst()
                for s_fname in l_flist :
#                    if os.path.splitext( s_fname )[1] == ".txt" :
#                        continue
                    
                    MsgFTP = FTP.retrlines("RETR " + s_fname, open(s_fname, 'wb').write)
                    if MsgFTP.find( "226" ) == -1 :
                        print "\t\t\t[Failure] RETRLines %s ( %s )" % ( s_fname, MsgFTP )
                        continue
                    
                    FileCount += 1
                
                if FileCount > 0 :
                    print "\t\t\t%s ( Count : %d )" %  ( s_extDir, FileCount)
                    CopyFlag = True
            
            if CopyFlag :
                return True
            else :
                return False
                        
        except :
            print format_exc()
            return False
        
        