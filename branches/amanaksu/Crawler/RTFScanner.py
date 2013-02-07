# -*- coding:utf-8 -*-

# import Public Module
import traceback

# import Private Module


class RTFScan():
    @classmethod
    def Check(cls, pBuf):
        try :
            
            RTFSig = pBuf[0:6]
            if RTFSig == "{\rtf1" or RTFSig.find( "rt" ) != -1 :
                return "RTF" 
            else :
                return None
            
        except :
            print  traceback.format_exc()
            return None