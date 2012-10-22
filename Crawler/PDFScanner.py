# -*- coding:utf-8 -*-

# import Public Module
import re, traceback


class PDFScan():
    @classmethod
    def Check(cls, pBuf):
        try : 
            if bool( re.match("^%PDF-[0-9]\.[0-9]", pBuf) ) :
                return True
            else :
                return False      
        except :
            print traceback.format_exc()