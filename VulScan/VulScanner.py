# -*- coding:utf-8 -*-

# import Public Module
import optparse, traceback, sys, os


# import Private Module
from ComFunc import FileControl
from PDFScanner import PDFScan
from OLEScanner import OLEScan



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
            File["pBuf"] = FileControl.ReadFileByBinary( File["fname"] )
            
            if File["format"] == "PDF" :
                PDFScan.Scan(File)
            elif File["format"] == "OLE" :
                OLEScan.Scan(File)
            else :
                return False
            
        except :
            print traceback.format_exc()
            
        return True



class Initialize():
    def __init__(self, File):
        # Init Input Options.....
        File['fpath'] = ""
        File['dpath'] = ""
        File['log'] = ""
        File['logbuf'] = "" 
        
        File['format'] = ""
        return
    
    
    def GetOption(self):
        Parser = optparse.OptionParser(usage='usage: %prog [-f|-d] file or folder\n')
        Parser.add_option('-f', '--file', help='< file name >')
        Parser.add_option('-d', '--directory', help='< directory >')
        Parser.add_option('--delete', help='< Ext Name>')
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
        
        FormatList = ["PDF", "OLE"]
        FormatFunc = {"PDF":PDFFile, "OLE":OLEFile}
        
        for field in FormatList :
            for fname in FormatFunc[ field ] :
                File["format"] = field
                File["fname"] = fname
                if not main.ScanFormat(File) :
                    File['logbuf'] += "\n[Failure] ScanFormat( %s )" % fname
                                
    except :
        print traceback.format_exc()
    
    Init.PrintLogOption(File)
    
    exit(0)
    
    
    