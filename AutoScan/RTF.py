# -*- coding:utf-8 -*-


from traceback import format_exc


class CRTF():
    @classmethod
    def fnCheck(cls, s_Buffer):
        try :
            RTFSig = s_Buffer[0:6]
            if RTFSig == "{\rtf1" or RTFSig.find( "rt" ) != -1 :
                return "RTF" 
            else :
                return None
            
        except :
            print format_exc()
            return None


    @classmethod
    def fnScan(cls, s_fname, s_pBuf, s_Format):
        
        print "\t[+] Scan RTF - %s" % s_fname
        
        try :
            
            return True
            
        except :
            print format_exc()
            return None