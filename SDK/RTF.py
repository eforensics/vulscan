# External Import
from traceback import format_exc


# Internal Import


class CRTF():
    @classmethod
    def fnIsRTF(cls, s_buffer):
        
        try :
            
            RTFSig = s_buffer[0:6]
            if RTFSig == "{\rtf1" or RTFSig.find( "rt" ) != -1 :
                return "RTF" 
            else :
                return None
            
        except :
            print format_exc()
            return None
    
    @classmethod
    def fnScanRTF(cls):
        pass
    

class CStructRTF():
    pass




class CPrintRTF():
    pass







