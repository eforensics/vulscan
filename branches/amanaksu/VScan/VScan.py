# External Module
import os, sys
from traceback import format_exc
from optparse import OptionParser, OptionGroup

# Internal Module
try :
    s_abs_curdir = os.path.abspath( os.path.curdir )
    s_path_SDK = s_abs_curdir + "\\SDK"
    s_path_Exploit = s_abs_curdir + "\\Exploit"
    
    sys.path.append(s_path_SDK)
    sys.path.append(s_path_Exploit)
    
    from VScanner import CScan
except :
    print "[ImportError] Put in Modules in the VScan.py File Folder"
    exit(-1)


def main():
    try :
        # Init & Options
        Init = CInit()
        options = Init.GetOptions()
        if options == None :
            return False
        
        # Scan
        if options.scan == True :
            if not CScan.fnScan(options) :
                return False
        
        # Other things....
        
    except :
        print format_exc()
        return False
    
    return True

class CInit():
    def PrintLog(self):
        print "-----------------------------------------------------------------"
        print "  VScan ver 1.0                   Copyright (c) 2013, Project1"
        print "-----------------------------------------------------------------"
        print ""
    def GetOptions(self):
        self.PrintLog() 
            
        # Default Option
        parser = OptionParser(usage="%prog -f <FILE> | -d <Folder> [Options] [Debug Options]", version="%prog 1.0")
        parser.add_option("-f", "--fname", metavar="FILE", type="string", help="Detect File Name")
        parser.add_option("-d", "--dir", metavar="DIR", type="string", help="Detect Directory Name")
        parser.add_option("-s", "--scan", action="store_true", dest="scan", help="Scan Vulnerability")
            
        # Debug Option
        group = OptionGroup(parser, "Debug Options")
        group.add_option("-l", "--log", action="store_true", help="Saved File Log")
        group.add_option("-e", "--extract", type="string", dest="extract", help="Extract Data")
        group.add_option("-c", "--classify", action="store_true", help="File Classifying")
        parser.add_option_group(group)
            
        (Options, args) = parser.parse_args()
            
        # Except Check Parameter
        if sys.argv.__len__() < 2 :
            parser.print_help()
            return None
            
        return Options

if __name__ == "__main__" :
    main()