# -*- coding:utf-8 -*-

# import Public Module
import traceback

# import Private Module
from ComFunc import BufferControl

class PEScan():
    @classmethod
    def Check(cls, pBuf):
        try :
            if BufferControl.Read(pBuf, 0, 2) == "MZ" or BufferControl.Read(pBuf, 0, 2) == "mz" :
                return "PE"
            else :
                return ""
            
        except :
            print traceback.format_exc()
            
        