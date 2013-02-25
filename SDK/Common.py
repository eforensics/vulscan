# External Import
import os
from struct import unpack
from traceback import format_exc

# Internal Import

class CFile():
    @classmethod
    def fnCreateFile(cls, s_fname, s_privilege):        
        p_file = None
        
        try :

            p_file = open(s_fname, "rb")
            
        except IOError :
            print "[-] Failed - IOError : %s" % s_fname
        
        except :
            print format_exc()
        
        return p_file
    @classmethod
    def fnReadFile(cls, s_fname):
        s_file = None
        
        try :
            
            with open(s_fname, "rb") as p_file :
                s_file = p_file.read()
        
        except IOError, error :
            print error
        
        except :
            print format_exc()
        
        return s_file
    @classmethod
    def fnWriteFile(cls, s_fname, s_file):
        try :
            with open(s_fname, "wb") as p_file :
                try :
                    p_file.write(s_file)
                except :
                    p_file.write(str(s_file))
        except :
            print format_exc()
            return False
        
        return True

class CBuffer():
    @classmethod
    def fnReadData(cls, s_buffer, n_Position, n_Size):
        
        try :
            if n_Size == 1 :
                s_Sig = "<B"
            elif n_Size == 2 :
                s_Sig = "<H"
            elif n_Size == 4 :
                s_Sig = "<L"
            else :
                print "[-] Error - fnReadData( %d )" % n_Size
                return None
            
            s_Val = s_buffer[n_Position:n_Position + n_Size]
            return unpack( s_Sig, s_Val )[0]
            
        except :
            print format_exc()
            return None
    @classmethod
    def fnBitParse(cls, n_Data, b_BitOffset, b_Cnt):
        
        try :
            
            b_Bit = (n_Data / 2**b_BitOffset) & (2**b_Cnt-1)
            return b_Bit
        
        except :
            print format_exc()
            return None


class CCheck():
    def fnCheckEnableFile(self, s_fname):
        h_file = None
        
        try :
            h_file = CFile.fnCreateFile(s_fname, "rb")
            if h_file != None :
                return True
                
        except IOError, error:
            print error
            
        except :
            print format_exc()
            
        return False






