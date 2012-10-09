# -*- coding:utf-8 -*-

# import Public Module
import traceback, binascii


# import private module
from ComFunc import BufferControl, FileControl


class Office():
    @classmethod
    def OfficeScan(cls, File):
        try :
            File['logbuf'] += " : Office"
        except :
            print traceback.format_exc()
            return False
        
        return True




