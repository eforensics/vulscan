# -*- coding:utf-8 -*-

# import Public Module
import traceback

# import Private Module
import Common



class PE():
    @classmethod
    def Check(cls, pBuf):
        try :
            if Common.BufferControl.Read(pBuf, 0, 2) == "MZ" or Common.BufferControl.Read(pBuf, 0, 2) == "mz" :
                return "PE"
            else :
                return ""
            
        except :
            print traceback.format_exc()
