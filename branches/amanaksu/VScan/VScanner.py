# External Module
import os
from traceback import format_exc

# SDK Module
try :
    from PDF import *
    from RTF import *
    from OLE import *
    from PE import *
    from Common import CFile
#    from Common import CFile
except :
    print "[-] Failed - File Import ( Format Module )"
    exit(-1)

# Internal Import Module



class CScan():
    l_Format = ["PDF", "RTF", "OLE", "PE", "Unknown"]
    l_Known = ["PDF", "RTF", "OLE", "PE"]
    
    l_Unknwon = []
    l_Except = []
    l_PDF = []
    l_RTF = []
    l_OLE = []
    l_OLE_HWP = []
    l_OLE_OFF = []
    l_PE = []    
    
    @classmethod
    def fnScan(cls, options):
        
        print "[+] Start"
        
        l_Files = []
        
        try :
            # Step 1. Create File List
            Scan = CScan()
            l_Files = Scan.fnCreateFileList(options)
            if l_Files == [] :
                print "[-] Error - fnCreateFileList()"
                return False
            
            # Step 2. Format Scanning
            for s_fname in l_Files :
                s_format = Scan.fnIsScannable(s_fname)
                eval( "cls.l_%s" % s_format ).append( s_fname )
            
            # Step 3. Mapped Scan
            ScanFile = CScanFile()
            for s_format in cls.l_Known :
                if not ScanFile.fnScanFileList(s_format, l_Files) :
                    print "[-] Error - fnScanFileList( %s )" % s_format
            
            # Step 4. Move Files ( .bin, .txt, .log )
            
            
            
        except :
            print format_exc()
            return False
        
        return True
    def fnCreateFileList(self, options):
        l_Files = []
        l_tmpFiles = []
        
        try :
            if options.fname != None :             
                l_Files.append( options.fname )
                
            if options.dir != None :
                l_tmpFiles = os.listdir( options.dir )
                for s_tmpFile in l_tmpFiles :
                    if os.path.isdir( s_tmpFile ) :
                        continue
                    l_Files.append( s_tmpFile )
            
        except :
            print format_exc()
        
        return l_Files
    def fnIsScannable(self, s_fname):
        s_flagbuffer = None        
        s_format = "unknown"
        
        try :
            
            s_flagbuffer = CFile.fnReadFile( s_fname )
            if s_flagbuffer == None :
                print "[-] Error - fnReadFile( %s )" % s_fname
                return s_format
            
            s_format = self.fnIsFormat(s_flagbuffer)
            
        except :
            print format_exc()
            return s_format
        
        return s_format
    def fnIsFormat(self, s_buffer):
        s_format = None

        try :
            
            for s_known in self.l_Known :
                s_format = eval("C%s.fnIs%s" % (s_known, s_known))(s_buffer)
                if s_format != None :
                    break
                
        except :
            print format_exc()
            return s_format
        
        return s_format
    def fnClassify(self):
        pass


class CScanFile():
    s_FileVersion = ""
    
    def fnScanFileList(self, s_format, l_files):
        
        try :
            
            Scan = CScan()
            for s_fname in l_files :
                if not eval( "self.fnScan%s" % s_format )( s_fname ) :
                    Scan.l_Except.append( s_fname )
            
        except :
            print format_exc()
            
        return True
    def fnScanPDF(self, s_fname):
        
        print "[+] PDF"
        
        try :
            
            return True
        
        except :
            print format_exc()
            
        return True
    def fnScanOLE(self, s_fname):
        
        print "[+] OLE",
        
        try :
            Scan = CScan()
            OLE = COLE()
            
            s_buffer = CFile.fnReadFile( s_fname )
            s_format = OLE.fnIsSubFormat( s_buffer )
            if s_format != "" :
                eval( "Scan.l_OLE_%s" % s_format ).append( s_fname )
                eval( "self.fnScan%s" % s_format )( s_buffer )
            else :
                return False
            
        except :
            print format_exc()
            
        return True
    def fnScanOffice(self, s_buffer):
        
        print "- Office"
        
        try :
            
            return True
        
        except:
            print format_exc()
            
        return True
    def fnScanHWP(self, s_buffer):
        
        print "- HWP"
        
        try :
            
            HWP = CHWP()
            HWP.fnFindHWPHeader(s_buffer, dl_OLEDirectory, t_SAT, t_SSAT)
            
            
            return True
        
        except :
            print format_exc()
            
        return True
    def fnScanRTF(self, s_fname):
        
        print "[+] RTF"
        
        try :
            
            return True
        
        except :
            print format_exc()
            
        return True        
    def fnScanPE(self, s_fname):
        
        print "[+] PE"
        
        try :
            
            return True
        
        except :
            print format_exc()
            
        return True
    

















    
    


























