# -*- coding:utf-8 -*-

from traceback import format_exc
from string import letters, digits
from struct import unpack

class CFile():
    
    #    s_fname                    [IN]        File Full Name
    #    p_file                     [OUT]       File Object
    @classmethod
    def fnOpenFile(cls, s_fname):
        try :
            p_file = open( s_fname, 'rb' )
            
        except IOError :
            print "[fnOpenFile] IOError : %s" % s_fname    
        
        except :
            print format_exc()
            
        return p_file
    
    
    #    s_fname                    [IN]        File Full Name
    #    s_pBuf                     [OUT]       File Buffer ( Type : Binary )
    @classmethod
    def fnReadFile(cls, s_fname):
        try :
            p_file = CFile.fnOpenFile(s_fname)     
            s_pBuf = p_file.read()
            p_file.close()
        
        except :
            print format_exc()
        
        return s_pBuf
    
    
    #    s_fname                    [IN]             File Full Name
    #    s_Buffer                   [IN]             Written Buffer Data
    #    BOOL Type                  [OUT]            True / False
    @classmethod
    def fnWriteFile(cls, s_fname, s_Buffer):
        try :
            with open(s_fname, "wb") as p_file :
                p_file.write( s_Buffer )
            
        except :
            print format_exc()
            return False
            
        return True



class CBuffer():
    
    #    s_pBuf                                [IN]                    Repair Buffer
    #    s_OutBuf                              [OUT]                   Repaired Buffer
    @classmethod
    def fnExtractAlphaNumber(cls, s_Buffer):
        try :
            s_Pattern = letters + digits
            l_ExtList = []
            for s_OneByte in s_Buffer :
                if s_OneByte in s_Pattern :
                    l_ExtList.append( s_OneByte )
            
            s_OutBuf = ''.join( l_ExtList )
            
        except :
            print format_exc() 
            return None

        return s_OutBuf

   
    #    s_Buffer                            [IN]                    File Buffer
    #    Position                            [IN]                    Reading Position
    #                                        [OUT]                   Unpacked 2Bytes Data
    @classmethod
    def fnReadWord(cls, s_Buffer, n_Position):
        try :
            s_Val = s_Buffer[n_Position:n_Position+2]
            return unpack("<H", s_Val)[0]
        
        except :
            print format_exc()
            return None
   
   
    #    s_Buffer                            [IN]                    File Buffer
    #    Position                            [IN]                    Reading Position
    #                                        [OUT]                   Unpacked 4Bytes Data
    @classmethod
    def fnReadDword(cls, s_Buffer, n_Position):
        try :
            s_Val = s_Buffer[n_Position:n_Position+4]
            return unpack("<L", s_Val)[0]
        
        except :
            print format_exc()
            return None
        
   
   
   





 
    
