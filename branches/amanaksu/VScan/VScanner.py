# External Module
import os, sys
from traceback import format_exc

# Internal Import Module
# SDK Module
try :
    from PDF import *
    from RTF import *
    from OLE import *
    from PE import *
    from Common import CFile
except :
    print "[ImportError] Do not Found SDK Modules"
    exit(-1)

# Exploit Module
try :   
    from ExploitHWP import CExploitHWP
    from ExploitOffice import CExploitOffice
    from ExploitRTF import CExploitRTF
    from ExploitPDF import CExploitPDF
except :
    print "[-] Warning - Do not Scan Exploit!"


class CScan():
    l_Format = ["PDF", "RTF", "OLE", "PE", "Unknown"]
    l_Known = ["PDF", "RTF", "OLE", "PE"]
    
    l_Unknwon = []
    l_Except = []
    l_PDF = []
    l_RTF = []
    l_OLE = []
    l_OLE_HWP = []
    l_OLE_Office = []
    l_PE = []    
    
    @classmethod
    def fnScan(cls, options):
        # [IN]     options
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] FALSE
        #    - [ERROR] None
        print "[+] Start"
        
        l_Files = []
        
        try :
#            print "Step 1. Create File List"
            Scan = CScan()
            l_Files = Scan.fnCreateFileList(options)
            if l_Files == None :
                print "[-] Error - fnCreateFileList()"
                return False
            elif l_Files == [] :
                print "[-] Failure - fnCreateFileList()"
                return False
            
#            print "Step 2. Format Scanning"
            for s_fname in l_Files :
                s_format = Scan.fnIsScannable(s_fname)
                if s_format in cls.l_Format :
                    eval( "cls.l_%s" % s_format ).append( s_fname )
            
#            print "Step 3. Mapped Scan"
            ScanFile = CScanFile()
            for s_format in cls.l_Known :
                if eval( "cls.l_%s" % s_format ) == [] :
                    continue
                b_Ret = ScanFile.fnScanFileList(s_format, options, eval( "cls.l_%s" % s_format ))
                if b_Ret == None :
                    print "[-] Error - fnScanFileList( %s )" % s_format
            
            # Step 4. Move Files ( .bin, .txt, .log )
            if options.classify == True :
                print "classify files by format"
            
        except :
            print format_exc()
            return False
        
        return True
    def fnCreateFileList(self, options):
        # [IN]    options
        # [OUT]   l_Files
        #    - [SUCCESS] file list in directory
        #    - [FAILURE] []
        #    - [ERROR] None
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
                os.chdir( options.dir )
                
        except :
            print format_exc()
            return None
        
        return l_Files
    def fnIsScannable(self, s_fname):
        # [IN]    s_fname
        # [OUT]   s_format
        #    - [SUCCESS] file format name
        #    - [FAILURE] "unknown"
        #    - [ERROR] "unknown"
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
        # [IN]    s_buffer
        # [OUT]   s_format
        #    - [SUCCESS] file format
        #    - [FAILURE/ERROR] None
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
    
    def fnScanFileList(self, s_format, options, l_files):
        # [IN]    s_format
        # [IN]    options
        # [IN]    l_files
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE/ERROR] None 
        try :
        
            Scan = CScan()
            for s_fname in l_files :
                if not eval( "self.fnScan%s" % s_format )( s_fname ) :
                    Scan.l_Except.append( s_fname )
                
                if options.log == True :
                    print "saved file scan log"
                    
                if options.extract == True :
                    print "extracted data"
            
        except :
            print format_exc()
            return None
            
        return True
    def fnScanPDF(self, s_fname):
        # [IN]    s_fname
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        print "[+] PDF :",
        
        try :
            print s_fname
            
            PDF = CPDF()
            PDF.s_fname = s_fname
            
            # Scan PDF Structure
            s_buffer = CFile.fnReadFile(s_fname)
            if not PDF.fnScanPDF(PDF, s_buffer) :
                print "[-] Error - fnScanPDF()"
                return False
            
            # Scan PDF Exploit
            #    - JavaScript 
            if PDF.l_JSIndex != [] :
                ExploitPDF = CExploitPDF()
                if not ExploitPDF.fnScanExploit(PDF) :
                    print "[-] Error - fnScanExploit()"
                    return False
            #    - Another Check ( Ex : Embedded PE, etc... )
            if PDF.l_JSIndex_Sus != [] :
                print "Another Suspicious Stream"
            # Print PDF
        
        except :
            print format_exc()
            return None
            
        return True
    def fnScanOLE(self, s_fname):
        # [IN]    s_fname
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        print "[+] OLE",
        
        try :
            # for Saved File List
            Scan = CScan()
            
            # for OLE's Function Instance
            OLE = COLE()
            
            s_buffer = CFile.fnReadFile( s_fname )
            if not OLE.fnScanOLE(OLE, s_buffer) :
                print "[-] Error - fnScanOLE()"
                return False
            
            s_format = OLE.fnIsSubFormat( OLE )
            if s_format != "" :
                eval( "Scan.l_OLE_%s" % s_format ).append( s_fname )
                b_Ret = eval( "self.fnScan%s" % s_format )(OLE, s_fname, s_buffer) 
                if b_Ret == False :
                    print "[-] Failure - Scan%s()" % s_format
                elif b_Ret == None :
                    print "[-] Error - Scan%s()" % s_format
                
            else :
                return False
            
            # for Print OLE Structure
#            PrintOLE = CPrintOLE()
#            PrintOLE.fnPrintOLEHeader()
#            PrintOLE.fnPrintOLEDirectory(OLE.dl_OLEDirectory)
            
        except :
            print format_exc()
            
        return True
    def fnScanOffice(self, OLE, s_fname, s_buffer):
        # [IN]    s_fname
        # [IN]    s_buffer
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        print "- Office :",
        
        try :
            
            print s_fname
            
            # Scan Office Structure
            Office = COffice()
            if not Office.fnScanOffice( OLE, Office, s_buffer ) :
                print "[-] Error - fnScanOffice()"
                return False
            
            # Scan Office Exploit
            ExploitOffice = CExploitOffice()
            if not ExploitOffice.fnScanExploit(OLE, Office, s_buffer) :
                return False
            
        except:
            print format_exc()
            return None
            
        return True
    def fnScanHWP(self, OLE, s_fname, s_buffer):
        # [IN]    s_fname
        # [IN]    s_buffer
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        print "- HWP :",
        
        try :
            
            print s_fname
            
            # Scan HWP Structure
            HWP = CHWP()
            if not HWP.fnScanHWP( OLE, s_buffer ) :
                print "[-] Error - fnScanHWP()"
                return False
            
            # Scan HWP Exploit
            ExploitHWP = CExploitHWP()
            if not ExploitHWP.fnScanExploit( HWP.l_HWPSSector ) :
                print "[-] Error - fnScanExploit()"
                return False
            
        except :
            print format_exc()
            
        return True
    def fnScanRTF(self, s_fname):
        # [IN]    s_fname
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        print "[+] RTF :",
        
        try :
            
            print s_fname
            return True
        
        except :
            print format_exc()
            
        return True        
    def fnScanPE(self, s_fname):
        # [IN]    s_fname
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        print "[+] PE :",
        
        try :
            
            print s_fname
            return True
        
        except :
            print format_exc()
            
        return True
    
