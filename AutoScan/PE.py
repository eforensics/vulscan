# -*- coding:utf-8 -*-


from traceback import format_exc


class CPE():
    @classmethod
    def Check(cls, s_Buffer):
        try :
            PESig = s_Buffer[0:2]
            if PESig == "MZ" or PESig == "mz" :
                return "PE" 
            else :
                return None
            
        except :
            print format_exc()
            return None
