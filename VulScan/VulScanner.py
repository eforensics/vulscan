# -*- coding:utf-8 -*-

# import Public Module
import optparse, traceback, sys, os


# import Private Module
from ComFunc import FileControl
from PDFScanner import PDFScan, PDFSearch, OperateStream
from OLEScanner import OLEScan, OLEStruct
from HWPScanner import HWP
from OfficeScanner import Office



class Main():
    def CheckFormat(self, fname):
        try :
            pBuf = FileControl.ReadFileByBinary(fname)
            if PDFScan.Check( pBuf ) :
                return "PDF"
            elif OLEScan.Check( pBuf ) :
                return "OLE"
            else :
                return ""
            
        except :
            print traceback.format_exc()
        
    
    def ScanFormat(self, File):
        try :
            scan = Scan()
                        
            File["pBuf"] = FileControl.ReadFileByBinary( File["fname"] )
            
            if File["format"] == "PDF" :
                scan.PDFScan(File)
#                PDFScan.Scan(File)
            elif File["format"] == "OLE" :
                scan.OLEScan(File)
#                OLEScan.Scan(File)
            else :
                return False
            
        except :
            print traceback.format_exc()
            
        return True


class Scan():
    def PDFScan(self, File):
        try :
            File['logbuf'] += "\n    [+] %s...........%s" % ( File["fname"], File["format"] )
            
#            if File["OptSep"] == "*" or File["OptSep"] == "PDF" :
#                if not FileControl.SeperateFile(File, "PDF") :
#                    File['logbuf'] += "\n\t[Failure] Move %s" % File["fname"]            
#            
#            return True
            
            pBuf = File["pBuf"]
            Search = PDFSearch()
            
            # Check Suspicious Events
            Event = Search.SearchEvent(pBuf)
            
            # Check Suspicious Actions
            Action = Search.SearchAction(pBuf)
            
            # Check Element to Analyze with Actions
            Element = Search.SearchElement(File, pBuf, Action)
            
            # Check UnEscape to Analyze Element in PDF     
            Operation = OperateStream()
            unescapedBytes = []
            for element in Element :
                if type(element) == type(None) :
                        continue
                
#                print "Scan : %s" % File["fname"]
                element = Operation.FilterObj(element)
                if Operation.isJavaScript( element ) :
                    name = File["fname"] + "_" + hex( len(element) )
                    FileControl.WriteFile(name, element)
                    
                    UnEscapeData = Operation.Unescaping( element )
                    if UnEscapeData not in unescapedBytes :
                        unescapedBytes.append( UnEscapeData )
            
            ElementData = ""
            for element in Element :
                JSList = Search.SearchJS(element)
                if JSList == [] :
                    continue

                for JS in JSList :
                    ElementData += "[%s : %s]\n" % (JS, CVENo[JS])
                ElementData += element
            
            
            if Event == [] and Action == [] and ElementData == "" :
                File['logbuf'] += "\t[ Done ]"
            else :
                File['logbuf'] += "\t[ Find ]"
            
            
            # Logging Suspicious Data 
            if Event != [] :
                File['logbuf'] += "\n            Event  : "
                for event in Event :
                    File['logbuf'] += event + " "
                
                
            if Action != [] :
                File['logbuf'] += "\n            Action : "
                for action in Action :
                    File['logbuf'] += action + " "
                
                
            if ElementData != "" :
                FileControl.WriteFile("%s_Element.dat" % File["fname"], ElementData)    
                File['logbuf'] += "\n        [-] %s\t<%s_Element.dat>" % (File["fname"],  File["fname"]) 
            
        except :
            File['logbuf'] += "\t[ Error ]"
            print traceback.format_exc()
            return False
        
        return True
    
    
    def OLEScan(self, File):
        try :
            File['logbuf'] += "\n    [+] %s...........%s" % ( File["fname"], File["format"] )
            
            OLE = OLEStruct( File )
            
            if not OLE.OLEHeader(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLEHeader( %s )" % File["fname"]
                return False
    
            
            if not OLE.OLETableSAT(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLETableSAT( %s )" % File["fname"]
                return False
    
            if not OLE.OLETableSSAT(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLETableSSAT( %s )" % File["fname"]
    
            
            if not OLE.OLEDirectory(File) :
                File['logbuf'] += "\n    [Failure] OLE.OLEDirectory( %s )" % File["fname"]
                return False
            
            
            DocScan = {"HWP":HWP.HWPScan, "Office":Office.OfficeScan}
            for field in DocScan :
                if File["format"] == field :
                    DocScan[ field ]( File ) 
            
        except :
            print traceback.format_exc()
            return False
        
        return True
     


#----------------------------------------------------------------------------------------------------------------------------------------------------#
#     for PDFScan
#----------------------------------------------------------------------------------------------------------------------------------------------------#
#SuspiciousEvents = ['/OpenAction', '/AA']
#SuspiciousActions = ['/JS', '/JavaScript', '/Launch', '/SubmitForm', '/ImportData']
#SuspiciousJS = ['mailto', 
#                'Collab.collectEmailInfo', 
#                'util.printf', 
#                'getAnnots', 
#                'getIcon', 
#                'spell.customDictionaryOpen', 
#                'media.newPlayer']

CVENo = {"mailto"                       :   "CVE-2007-5020", 
         "Collab.collectEmailInfo"      :   "CVE-2007-5659", 
         "util.printf"                  :   "CVE-2008-2992", 
         "getAnnots"                    :   "CVE-2009-1492", 
         "getIcon"                      :   "CVE-2009-0927", 
         "spell.customDictionaryOpen"   :   "CVE-2009-1493", 
         "media.newPlayer"              :   "CVE-2009-4324"}






class Initialize():
    def __init__(self, File):
        # Init Input Options.....
        File['fpath'] = ""
        File['dpath'] = ""
        File['log'] = ""
        File['logbuf'] = "" 
        
        File['format'] = ""
        
        File['OptSep'] = ""
        
        return
    
    
    def GetOption(self):
        Parser = optparse.OptionParser(usage='usage: %prog [-f|-d] file or folder\n')
        Parser.add_option('-f', '--file', help='< file name >')
        Parser.add_option('-d', '--directory', help='< directory >')
        Parser.add_option('--separate', help='< * | Format Extend >')
        Parser.add_option('--delete', help='< Ext Name >')
        Parser.add_option('--log', help='< file name >')
        (options, args) = Parser.parse_args()
            
        if len( sys.argv ) < 2 :
            Parser.print_help()
            exit(0)
    
        return options
    
    
    def SetOption(self, File, Opt):
        try :
            if Opt.directory :
                File['dpath'] = Opt.directory
            
            if Opt.file :
                File['fpath'] = Opt.file
            
            if Opt.separate :
                File['OptSep'] = Opt.separate
            
            if Opt.delete :
                File['ExtName'] = Opt.delete
            
            if Opt.log : 
                File['log'] = Opt.log
            
        except :
            print traceback.format_exc()
            return False
        
        return True
        

    def PrintLogOption(self, File):
        try :
            if File['log'] == "" :
                print File['logbuf']
                
            else :
                FileControl.WriteFile( File['log'], File['logbuf'] )
        
        except :
            print traceback.format_exc()
            


if __name__ == '__main__' :
    
    File = {}
    Init = Initialize( File )
    
    Options = Init.GetOption()
    
    try :
        if not Init.SetOption( File, Options ) :
            File['logbuf'] += "\n[Failure] SetOption()"
            exit(-1)
        
        flist = []
        if Options.file :
            flist.append( Options.file )
        elif Options.directory :
            flist = os.listdir( Options.directory )
            os.chdir( Options.directory )            
        else :
            File['logbuf'] += "\n[Failure] Set FileList"
            exit(-1)
        
        
        # Option : Delete Files 
        if Options.delete :
            for fname in flist :
                fext = os.path.splitext( fname )
                fdel = os.path.splitext( Options.delete )
                
                if fext[1] == fdel[1] :
                    os.remove( fname )
        
            exit(0)
        
        
        File['logbuf'] += "[*] Vulnerability Scanner"
        
        # Check File Format
        File['logbuf'] += "\n[*] File Format Checking"
        PDFFile = []    
        OLEFile = []
        for fname in flist :
            main = Main()
            
            strFormat = main.CheckFormat(fname)
            if strFormat == "PDF" :
                PDFFile.append( fname )
            elif strFormat == "OLE" :
                OLEFile.append( fname )
            else :
                File['logbuf'] += "\n    [-] Not Support! : %s" % fname
        
        if PDFFile == [] and OLEFile == [] :
            File['logbuf'] += "\n[Failure] None for Scan File"
            exit(-1)     
        
        
        # Scan File 
        File['logbuf'] += "\n[*] File Scanning"
        
        File["PDFList"] = PDFFile
        File["OLEList"] = OLEFile
        
        main.ScanFormat(File)
            
        
        
#        FormatList = ["PDF", "OLE"]
#        FormatFunc = {"PDF":PDFFile, "OLE":OLEFile}
#        
#        for field in FormatList :
#            for name in FormatFunc[ field ] :
#                File["format"] = field
#                File["fname"] = name
#                if not main.ScanFormat(File) :
#                    File['logbuf'] += "\n[Failure] ScanFormat( %s )" % name
                                
    except :
        print traceback.format_exc()
    
    Init.PrintLogOption(File)
    
    exit(0)
    
    
    