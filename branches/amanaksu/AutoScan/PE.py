# -*- coding:utf-8 -*-


from traceback import format_exc


class CPE():
    @classmethod
    def fnCheck(cls, s_Buffer):
        try :
            PESig = s_Buffer[0:2]
            if PESig == "MZ" or PESig == "mz" :
                return "PE" 
            else :
                return None
            
        except :
            print format_exc()
            return None


    @classmethod
    def fnScan(cls, s_fname, s_pBuf, s_Format):
        
        print "\t[+] Scan PE - %s" % s_fname
        
        try :
            
            return True
        
        except :
            print format_exc()
            return None
            