# -*- coding:utf-8 -*-

# import Public Module
import struct, traceback, glob, os, string


# Related File Read/Write
# Property : ClassMethod
class BufferControl():
    
    @classmethod
    def Read(cls, pBuf, Position, Size):
        return pBuf[Position:Position+Size]
    
    @classmethod
    def ReadWord(cls, pBuf, Position):
        return struct.unpack("<h", pBuf[Position:Position+2])[0]
    
    @classmethod
    def ReadDword(cls, pBuf, Position):
        return struct.unpack("<L", pBuf[Position:Position+4])[0]
    
    @classmethod
    def ReadSectorByBuffer(cls, pBuf, SecID, Size):
        Sec_Data = ""
        Sec_Position = ( SecID + 1 ) * Size
        
        try :
            Sec_Data = BufferControl.Read(pBuf, Sec_Position, Size)
        except :
            print traceback.format_exc()
        
        return Sec_Data
    
    @classmethod
    def ReadSectorByPointer(cls, fp, SecID, Size):
        Sec_Data = ""
        Sec_Position = ( SecID + 1 ) * Size
        
        try :
            fp.seek( Sec_Position )
            Sec_Data = fp.read( Size )
        except :
            print traceback.format_exc()
        
        return Sec_Data
    
    @classmethod
    def ConvertBinary2List(cls, pBuf, Length):
        Buf_List = []
        index = 0
        Position = 0
        
        try : 
            while index < ( len(pBuf) / Length ) :
                if Length == 2 :
                    Buf_List.append( BufferControl.ReadWord(pBuf, Position) )
                elif Length == 4 :
                    Buf_List.append( BufferControl.ReadDword(pBuf, Position) )
                else :
                    print "[Except] Add Case in ConvertBinary2List( )"
                
                Position += Length
                index += 1
        except :
            print traceback.format_exc()
        
        return Buf_List 
    
    
    @classmethod
    def ExtractAlphaNumber(cls, pBuf):
        try :
            Pattern = string.letters + string.digits
            ExtList = []
            for OneByte in pBuf :
                if ( OneByte in Pattern ) :
                    ExtList.append( OneByte )
            
            OutBuf = ''.join( ExtList )
            return OutBuf
        except :
            print traceback.format_exc() 
    
    
    
class FileControl():
    @classmethod
    def OpenFileByBinary(cls, fname):
        try : 
            fp = open( fname, 'rb' )
        except :
            print traceback.format_exc()
            
        return fp
    
    
    @classmethod
    def WriteFile(cls, fname, Buf):
        try : 
            with open( fname, 'w' ) as fp :
                fp.write( Buf )
        except :
            print traceback.format_exc()
    
    @classmethod
    def DeleteFile(cls, ext):
        try : 
            flist = glob.glob( ext )
            for f in flist :
                os.remove(f)
        except :
            print traceback.format_exc()
    
    


