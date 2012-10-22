# -*- coding:utf-8 -*-

# import Public Module
import traceback

# import Private Module
from ComFunc import BufferControl

class OLEScan():
    @classmethod
    def Check(cls, pBuf):
        try :
            # Case1. 0xe011cfd0, 0xe11ab1a1L
            if BufferControl.ReadDword(pBuf, 0) == 0xe011cfd0L and BufferControl.ReadDword(pBuf, 0x4) == 0xe11ab1a1L :
                return True
            # Case2. 0xe011cfd0, 0x20203fa1
            elif BufferControl.ReadDword(pBuf, 0) == 0xe011cfd0L :
                return True
            else :
                return False
        
        except :
            print traceback.format_exc()