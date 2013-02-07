# -*- coding:utf-8 -*-

# import Public Module
import re, traceback, binascii


# import private module
from ComFunc import DecodeControl, FileControl


class PDFScan():
    @classmethod
    def Check(cls, pBuf):
        try : 
            if bool( re.match("^%PDF-[0-9]\.[0-9]", pBuf) ) :
                return "PDF"
            else :
                return ""      
        except :
            print traceback.format_exc()
        
        








