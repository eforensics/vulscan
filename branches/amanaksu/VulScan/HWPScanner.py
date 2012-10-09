# -*- coding:utf-8 -*-

# import Public Module
import traceback, binascii


# import private module
from ComFunc import BufferControl, FileControl


class HWP():
    @classmethod
    def HWPScan(cls, File):
        try :
            File['logbuf'] += " : HWP"
        except :
            print traceback.format_exc()
            return False
        
        return True





