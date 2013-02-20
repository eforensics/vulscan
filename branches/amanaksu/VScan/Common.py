# External Import
import os
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
    pass







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






