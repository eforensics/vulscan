# External Import
import os
from struct import unpack
from traceback import format_exc

# Internal Import

class CFile():
    @classmethod
    def fnCreateFile(cls, s_fname, s_privilege):     
        # [IN]    s_fname
        # [IN]    s_privilege
        # [OUT]   p_file
        #    - [SUCCESS] file pointer
        #    - [FAILURE/ERROR] None
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
        # [IN]    s_fname
        # [OUT]   s_file
        #    - [SUCCESS] file buffer
        #    - [FAILURE/ERROR] None
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
        # [IN]    s_fname
        # [IN]    s_file
        # [OUT]   
        #    - [SUCCESS] saved file
        #    - [FAILURE] don't saved file
        #    - [ERROR] False
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
        # [IN]    s_buffer
        # [IN]    n_Position
        # [IN]    n_Size
        # [OUT]
        #    - [SUCCESS] data referred n_Size
        #    - [FAILURE/ERROR] None
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




