# External Import Module
from traceback import format_exc

# Internal Import Module


# Define
SIZE_READ_LINE = 100


class CFile():
    @classmethod
    def fnCreateFile(cls, s_fname, s_privilege):   
        # [IN]    s_fname
        # [IN]    s_privilege
        # [OUT]   p_file
        #    - [SUCCESS] p_file = Not None
        #    - [FAILURE] p_file = None
        p_file = None
        
        try :

            p_file = open(s_fname, s_privilege)
            
        except :
            print format_exc()
        
        return p_file
    @classmethod
    def fnReadFile(cls, s_fname):
        # [IN]     s_fname
        # [OUT]    s_file
        #    - [SUCCESS] s_file : Not None
        #    - [FAILURE] s_file : None
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
        # [OUT]   Boolean
        #    - [SUCCESS] : True
        #    - [FAILURE] : False
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
    def fnReadLines(cls, p_file):
        # [IN]    p_file
        # [OUT]   s_Lines
        #    - [SUCCESS] s_Lines : String Data
        #    - [FAILURE] s_Lines : []
        l_Lines = []

        try :
            
            l_Lines = p_file.readlines(SIZE_READ_LINE)
            
        except :
            print format_exc()
            
        return l_Lines






