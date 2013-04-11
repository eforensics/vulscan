# External Import
import os
from struct import unpack
from traceback import format_exc

# Internal Import

class CFile():
    @classmethod
    def fnCreateFile(cls, s_fname, s_privilege):     
        # [IN]    s_fname
        # [IN]    s_privilege
        # [OUT]   p_file
        #    - [SUCCESS] file pointer
        #    - [FAILURE/ERROR] None
        p_file = None
        
        try :

            p_file = open(s_fname, "rb")
            
        except IOError :
            print "[-] Failed - IOError : %s" % s_fname
        
        except :
            print format_exc()
        
        return p_file
    @classmethod
    def fnReadFile(cls, s_fname):
        # [IN]    s_fname
        # [OUT]   s_file
        #    - [SUCCESS] file buffer
        #    - [FAILURE/ERROR] None
        s_file = None
        
        try :
            
            with open(s_fname, "rb") as p_file :
                s_file = p_file.read()
        
        except IOError, error :
            print error
        
        except :
            print format_exc()
        
        return s_file
    @classmethod
    def fnWriteFile(cls, s_fname, s_file):
        # [IN]    s_fname
        # [IN]    s_file
        # [OUT]   
        #    - [SUCCESS] saved file
        #    - [FAILURE] don't saved file
        #    - [ERROR] False
        try :
            with open(s_fname, "wb") as p_file :
                try :
                    p_file.write(s_file)
                except :
                    p_file.write(str(s_file))
        except :
            print format_exc()
            return False
        
        return True

class CBuffer():
    @classmethod
    def fnReadData(cls, s_buffer, n_Position, n_Size):
        # [IN]    s_buffer
        # [IN]    n_Position
        # [IN]    n_Size
        # [OUT]
        #    - [SUCCESS] data referred n_Size
        #    - [FAILURE/ERROR] None
        try :
            if n_Size == 1 :
                s_Sig = "<B"
            elif n_Size == 2 :
                s_Sig = "<H"
            elif n_Size == 4 :
                s_Sig = "<L"
            else :
                print "[-] Error - fnReadData( %d )" % n_Size
                return None
            
            s_Val = s_buffer[n_Position:n_Position + n_Size]
            return unpack( s_Sig, s_Val )[0]
            
        except :
            print format_exc()
            return None
    @classmethod
    def fnBitParse(cls, n_Data, b_BitOffset, b_Cnt):
        # [IN] n_Data
        # [IN] b_BitOffset
        # [IN] b_Cnt
        # [OUT] b_Bit
        #    - [SUCCESS] bit data referred b_BitOffset
        #    - [FAILURE/ERROR] None
        try :
            
            b_Bit = (n_Data / 2**b_BitOffset) & (2**b_Cnt-1)
            return b_Bit
        
        except :
            print format_exc()
            return None

class CUtil():
    @classmethod
    def fnStringOffsetFindAll(cls, s_buffer, s_sig):
        # [IN]    s_buffer
        # [OUT]   l_Index
        #    - [FIND] index list
        #    - [Non-FIND] []
        #    - [ERROR] None
        try :
            l_Index = []
            
            s_tmpbuffer = s_buffer
            while True :
                n_Index = -1
                n_Index = s_tmpbuffer.find( s_sig )
                if n_Index == -1 :
                    break
            
                l_Index.append( n_Index + (l_Index.__len__() * s_sig.__len__()) )
                if n_Index + s_sig.__len__() < s_tmpbuffer.__len__() :
                    s_tmpbuffer = s_tmpbuffer[n_Index + s_sig.__len__():]
                else :
                    break
            
        except :
            print format_exc()
            return None
        
        return l_Index
    @classmethod
    def fnUtilTakeOffString(cls, s_buffer, s_sig):
        # [IN]    s_buffer
        # [OUT]   s_Outbuffer
        #    - [SUCCESS] string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_Outbuffer = [y for y in s_buffer if not (y in s_sig)]
            
        except :
            print format_exc()
            return None
        
        return s_Outbuffer   
    @classmethod
    def fnIsJavaScript(cls, s_buffer):
        # [IN]    s_buffer
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE/ERROR] False
        try :
            
            l_JSStrings = ['var ',';',')','(','function ','=','{','}','if ','else','return','while ','for ',',','eval']
            l_KeyStrings = [';','(',')']
            l_FoundStrings = []
            n_limit = 15
            n_result = 0

            for s_char in s_buffer :
                if ord(s_char) >= 150 :
                    return False
                
            for s_JSString in l_JSStrings :
                n_cnt = s_buffer.count( s_JSString )
                n_result += n_cnt
                if n_cnt > 0 :
                    l_FoundStrings.append(s_JSString)
                elif n_cnt == 0 and s_JSString in l_KeyStrings :
                    return False
            
            if n_result > n_limit and l_FoundStrings.__len__() >= 5 :
                return True
            else :
                return False
            
        except :
            print format_exc()
            return False
        






