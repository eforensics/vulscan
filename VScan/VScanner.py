# External Module
import os
from traceback import format_exc

# SDK Module
try :
    from PDF import *
    from RTF import *
    from OLE import *
    from PE import *
except :
    print "[-] Failed - File Import ( Format Module )"
    exit(-1)

# Internal Module
from Common import CFile


class CScan():
    l_Format = ["PDF", "RTF", "OLE", "PE", "Unknown"]
    l_Known = ["PDF", "RTF", "OLE", "PE"]
    
    l_Unknwon = []
    l_PDF = []
    l_RTF = []
    l_OLE = []
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
            
            
            
            # Step 4. Move Files ( Scanned File, .txt, .log )
            
            
            
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
            
            s_flagbuffer = CFile.fnReadFile(s_fname)[:0x20]
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
                s_format = eval("fnIs%s" % (s_known, s_known))(s_buffer)
                if s_format != None :
                    break
                
        except :
            print format_exc()
            return s_format
        
        return s_format

    def fnClassify(self):
        pass


class CScanFile():
    def fnScanFileList(self, s_format, l_files):
        pass
    
    def fnScanPDF(self, s_fname):
        pass
    
    def fnScanRTF(self, s_fname):
        pass
    
    def fnScanOLE(self, s_fname):
        pass
    
    def fnScanPE(self, s_fname):
        pass
    

















    
    


























