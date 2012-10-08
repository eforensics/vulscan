# -*- coding:utf-8 -*-

# import Public Module
import optparse, sys, os, re, traceback, zlib, binascii

# import Private Module
from ComFunc import DecodeControl, FileControl



class Initial():
    def GetOption(self):
        Parser = optparse.OptionParser(usage='usage: %prog [-f|-d] file or folder\n')
        Parser.add_option('-f', '--file', help='< file name >')
        Parser.add_option('-d', '--directory', help='< directory >')
        Parser.add_option('--delete', help='< extend >')
        (options, args) = Parser.parse_args()
            
        if len( sys.argv ) < 2 :
            Parser.print_help()
            exit(0)
    
        return options


class AnalyObj():
    def ExtractObj(self, fname, pBuf):
        try : 
            TargetList = []
            
            FlateDecode = []
            ASCII85Decode = []
            ASCIIHexDecode = []
            RunLengthDecode = []
            LZWDecode = []
            NoneDecode = []
            
            # Set Decode List
            streams = re.findall('obj\s{0,10}<<(.{0,100}?/Length\s([0-9]{1,8}).*?)>>[\s%]*?stream(.*?)endstream[^=]',pBuf,re.DOTALL)
            for streamDict in streams :         
                fstream = "%s_Enc_%s.dat" % (fname, streamDict[1])                
                FileControl.WriteFile( fstream, streamDict[2])
                if streamDict[0].find('/FlateDecode') != -1 or streamDict[0].find('/Fl') != -1:
                    FlateDecode.append( fstream )
                elif streamDict[0].find('/ASCII85Decode') != -1 :
                    ASCII85Decode.append( fstream )
                elif streamDict[0].find('/ASCIIHexDecode') != -1 :
                    ASCIIHexDecode.append( fstream )
                elif streamDict[0].find('/RunLengthDecode') != -1 :
                    RunLengthDecode.append( fstream )
                elif streamDict[0].find('/LZWDecode') != -1 :
                    LZWDecode.append( fstream )
                else :
                    NoneDecode.append( fstream )        
            
            # Check Decode List
            TargetList.append( FlateDecode )
            TargetList.append( ASCII85Decode )
            TargetList.append( ASCIIHexDecode )
            TargetList.append( RunLengthDecode )
            TargetList.append( LZWDecode )
            TargetList.append( NoneDecode )             
        
        except :
            print traceback.format_exc()        
            return TargetList
        
        return TargetList
    
    
    def DecObj(self, dlist, index):
        try :  
            FControl = FileControl()
            for fname in dlist :
                print "[Decoding]  %s" % fname
                
                pBuf = ""
                DecStream = ""
                pBuf = FControl.ReadFileByBinary( fname )
                DecStream = DecodeFunc[ TargetDecode[index] ]( pBuf )            
                if DecStream != None :
                    FControl.WriteFile( "%s_Dec" % fname, DecStream )
               
        except :
            print traceback.format_exc()
            return False
        
        return True

DecodeFunc = {"FlateDecode":DecodeControl.FlateDecode, 
              "ASCII85Decode":DecodeControl.ASCII85Decode, 
              "ASCIIHexDecode":DecodeControl.ASCIIHexDecode, 
              "RunLengthDecode":DecodeControl.RunLengthDecode, 
              "LZWDecode":DecodeControl.LZWDecode, 
              "NoneDecode":None}
TargetDecode = ["FlateDecode", "ASCII85Decode", "ASCIIHexDecode", "RunLengthDecode", "LZWDecode", "NoneDecode"]



if __name__ == '__main__' :
    
    Init = Initial()
    Opt = Init.GetOption()

    try :
        flist = []        
        if Opt.file :
            fpath = os.path.split( Opt.file )
            # fpath[0] : Directory path
            # fpath[1] : file name 
            os.chdir( fpath[0] )
            flist.append( fpath[1] )
        elif Opt.directory :
            os.chdir( Opt.directory )
            flist = os.listdir( Opt.directory )          
        else :
            print "[ERROR] Set File List"
            exit(-1)
    
        if flist == [] :
            print "[END] File is NONE"
            exit(0)
       
        
        # Option : Delete Files 
        if Opt.delete :
            for fname in flist :
                fext = os.path.splitext( fname )
                fdel = os.path.splitext( Opt.delete )
                
                if fext[1] == fdel[1] :
                    os.remove( fname )
        
            exit(0)
        
        
        # Default : Scan Files for Extract & Analysis
        for fname in flist :
            Analysis = AnalyObj()
            pBuf = FileControl.ReadFileByBinary( fname )
            
            # Extract Object in file
            DecodeList = []
            DecodeList = Analysis.ExtractObj( fname, pBuf )     
            
#            print "FlateDecode : ",
#            print DecodeList[0]
#            print "ASCII85Decode : ",
#            print DecodeList[1]
#            print "ASCIIHexDecode : ",  
#            print DecodeList[2]
#            print "RunLengthDecode : ",
#            print DecodeList[3]
#            print "LZWDecode : ",
#            print DecodeList[4]
#            print "NoneDecode : ",
#            print DecodeList[5]   
            
            # Analysis Extracted Object
            index = 0
            for dlist in DecodeList :
                if dlist != [] :
                    # Except NoneDecode
                    if index == 5 :
                        continue
                    
                    if not Analysis.DecObj( dlist, index ) :
                        print "[ERROR] Decompressed Object ( %d )" % index
                        break
                index += 1
    
    except SystemExit :
        exit(0)        
    
    except :
        print traceback.format_exc()
    
    exit(0)



