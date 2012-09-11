# -*- coding:utf-8 -*-

# import Public Module
import struct, traceback


class CTRL():
    def __init__(self):
        return
    
    # Default Reading is Big-Endian
    @classmethod
    def Read(cls, fbuf, offset, size):
        return fbuf[offset:offset+size]
    
    # Default Reading is Little-Endian
    # struct.unpack format ( Referred Site : http://pydoc.org/2.4.1/struct.html )
    #    < : Little-Endian
    #    > : Big-Endian
    #    Lower : signed / Upper : unsigned
    #        c : char
    #        b : byte
    #        h : short
    #        i : int
    #        l : long
    #        f : float ( None Upper )
    #        d : double ( None Upper )
    @classmethod
    def Word(cls, fbuf, offset):
        return struct.unpack("<h", fbuf[offset:offset+2])[0]
    
    # Default Reading is Little-Endian
    @classmethod
    def Dword(cls, fbuf, offset):
        return struct.unpack("<L", fbuf[offset:offset+4])[0]
    
    
    @classmethod
    def ReadSector(cls, fp, SecID):
        Sec_Data = ""
        Sec_Position = (SecID + 1) * szSector
        
        try :
            fp.seek( Sec_Position )
            Sec_Data = fp.read( szSector )
        except :
            print traceback.format_exc()
            
        return Sec_Data
    
    
    @classmethod
    def Bin2List(cls, buf, length):
        buf_list = []
        offset = 0
        index = 0
        
        while index < ( len(buf) / length ) :
            if length == 2 :
                buf_list.append( CTRL.Word(buf, offset) )
            elif length == 4 :
                buf_list.append( CTRL.Dword(buf, offset) )
            else :
                print traceback.format_exc()
            
            offset += length
            index += 1
        
        return buf_list
        
    
    
    
    

szSector = 0x200