# External Import
import re
from traceback import format_exc


# Internal Import



class CPDF():
    @classmethod
    def fnIsPDF(cls, s_buffer):
        
        try :
            
            if bool( re.match("^%PDF-[0-9]\.[0-9]", s_buffer) ) :
                return "PDF"
            else :
                return None
            
        except :
            print format_exc()
            return None
    
    @classmethod
    def fnScanPDF(cls):
        pass


class CStructPDF():
    pass



class CPrintPDF():
    pass




