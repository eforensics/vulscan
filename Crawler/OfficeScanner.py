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
            
#            if File["OptSep"] == "*" or File["OptSep"] == "Office" :
#                if not FileControl.SeperateFile(File, "Office") :
#                    File['logbuf'] += "\n\t[Failure] Move %s" % File["fname"]            
#            
#            return True
            
        except :
            print traceback.format_exc()
            return False
        
        return True




