# External Import
from traceback import format_exc


# Internal Import

class CPE():
    @classmethod
    def fnIsPE(cls, s_buffer):
        
        try :
            
            PESig = s_buffer[0:2]
            if PESig == "MZ" or PESig == "mz" :
                return "PE" 
            else :
                return None
            
        except :
            print format_exc()
            return None
    
    @classmethod
    def fnScanPE(cls):
        pass