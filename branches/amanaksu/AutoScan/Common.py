# -*- coding:utf-8 -*-




from traceback import format_exc
from string import letters, digits
from struct import unpack
from collections import namedtuple


class CFile():
    
    #    s_fname                    [IN]        File Full Name
    #    p_file                     [OUT]       File Object
    @classmethod
    def fnOpenFile(cls, s_fname):
        try :
            p_file = open( s_fname, 'rb' )
            
        except IOError :
            print "[fnOpenFile] IOError : %s" % s_fname    
        
        except :
            print format_exc()
            
        return p_file
    
    
    #    s_fname                    [IN]        File Full Name
    #    s_pBuf                     [OUT]       File Buffer ( Type : Binary )
    @classmethod
    def fnReadFile(cls, s_fname):
        try :
            p_file = CFile.fnOpenFile(s_fname)     
            s_pBuf = p_file.read()
            p_file.close()
        
        except :
            print format_exc()
        
        return s_pBuf
    
    
    #    s_fname                    [IN]             File Full Name
    #    s_Buffer                   [IN]             Written Buffer Data
    #    BOOL Type                  [OUT]            True / False
    @classmethod
    def fnWriteFile(cls, s_fname, s_Buffer):
        try :
            with open(s_fname, "wb") as p_file :
                p_file.write( s_Buffer )
            
        except :
            print format_exc()
            return False
            
        return True



class CBuffer():
    
    #    s_pBuf                                [IN]                    Repair Buffer
    #    s_OutBuf                              [OUT]                   Repaired Buffer
    @classmethod
    def fnExtractAlphaNumber(cls, s_Buffer):
        try :
            s_Pattern = letters + digits
            l_ExtList = []
            for s_OneByte in s_Buffer :
                if s_OneByte in s_Pattern :
                    l_ExtList.append( s_OneByte )
            
            s_OutBuf = ''.join( l_ExtList )
            
        except :
            print format_exc() 
            return None

        return s_OutBuf


    #    s_Buffer                            [IN]                    File Buffer
    #    Position                            [IN]                    Reading Position
    #                                        [OUT]                   Unpacked 1Byte Data
    @classmethod
    def fnReadByte(cls, s_Buffer, n_Position):
        try :
            s_Val = s_Buffer[n_Position:n_Position+1]
            return unpack("<B", s_Val)[0]
        
        except :
            print format_exc()
            return None

   
    #    s_Buffer                            [IN]                    File Buffer
    #    Position                            [IN]                    Reading Position
    #                                        [OUT]                   Unpacked 2Bytes tuple  Data
    @classmethod
    def fnReadWord(cls, s_Buffer, n_Position):
        try :
            
            s_Val = s_Buffer[n_Position:n_Position+2]
            return unpack("<H", s_Val)[0]
        
        except :
            print format_exc()
            return None
   
   
    #    s_Buffer                            [IN]                    File Buffer
    #    Position                            [IN]                    Reading Position
    #                                        [OUT]                   Unpacked 4Bytes Data
    @classmethod
    def fnReadDword(cls, s_Buffer, n_Position):
        try :
            
            s_Val = s_Buffer[n_Position:n_Position+4]
            return unpack("<L", s_Val)[0]
        
        except :
            print format_exc()
            return None
    
    
    #    s_Buffer                            [IN]                    String Buffer
    #    s_Seperator                         [IN]                    Seperator
    #    l_Buffer                            [OUT]                   String List By Seperator
    @classmethod
    def fnStr2List(cls, s_Buffer, s_Seperator):
        
        l_Buffer = []
        
        try :
            
            s_Data = ""
            for s_tmpData in s_Buffer :
                if s_tmpData == s_Seperator :
                    l_Buffer.append( s_Data )
                    s_Data = ""
                    continue
                
                s_Data += s_tmpData
                
            if s_Data != "" :
                l_Buffer.append( s_Data )
            
            return l_Buffer
        
        except :
            print format_exc()
            return None
    
   
    #    n_Data                                    [IN]                        Parse Target Data ( Little-Endian )
    #    b_BitOffset                               [IN]                        Parse Target Bit
    #            Ex)        0x1200
    #                       ========================================
    #               Offset  0 1 2 3  4 5 6 7  8 9 10 11  12 13 14 15
    #                       ----------------------------------------
    #               Data    0 0 0 0  0 0 0 0  1 0  0  0   0  1  0  0
    #                       ========================================
    #    b_Binary                                  [IN]                        Operate Bit ( 0 or 1 )
    #    b_Bit                                     [OUT]                       Operated Bit ( 0 or 1 )
    @classmethod 
    def fnBitParse(cls, n_Data, b_BitOffset, b_Binary):
        
        try :
            
            b_Bit = (n_Data / 2**b_BitOffset) & (2**b_Binary-1)
            return b_Bit
            
        except :
            print format_exc()
            return None


    #    s_Rule                                    [IN]                        Parse Pattern Rule String
    #    n_Data                                    [IN]                        Parse Target Data ( Little-Endian )
    #    l_Unpack                                  [OUT]                       Unpacked Data 
    @classmethod
    def fnBitUnpack(cls, s_Rule, n_Data):
        
        try :
            l_Unpack = []
            b_BitCount = 0
            s_Pattern = s_Rule.split(',')
            for b_Size in s_Pattern :
                if b_Size == "?" :
                    b_tmpSize = CBuffer.fnBitCount(n_Data)
                    if b_tmpSize != None :
                        b_Size = b_tmpSize
                else :
                    b_Size = int( b_Size )
    
                l_Unpack.append( (n_Data / 2 ** b_BitCount) & (2 ** b_Size-1) )
                b_BitCount += int( b_Size )
        
        except :
            print format_exc()
            return None
        
        return l_Unpack
    
    
    #    s_Buffer                                    [IN]                        Mapped Buffer
    #    s_NameString                                [IN]                        Mapped Name String
    #    s_NamePattern                               [IN]                        Mapped Name Pattern
    #    s_DataPattern                               [IN]                        Mapped Data Pattern
    #    s_Unpack                                    [OUT]                       Unpacked Data
    @classmethod
    def fnUnpack(cls, s_Buffer, s_NameString, s_NamePattern, s_DataPattern):
        
        try :
            
            t_Name = namedtuple( s_NameString, s_NamePattern )
            s_Unpack = t_Name._make( unpack(s_DataPattern, s_Buffer) )
            return s_Unpack
            
        except :
            print format_exc()
            return None
        
    
    
    #    n_Data                                    [IN]                        Parse Target Data ( Little-Endian )
    #    n_BitCount                                [OUT]                       Bit Count
    @classmethod
    def fnBitCount(cls, n_Data):

        try :
        
            n_BitCount = 0
            while True :
                n_Data = n_Data / 2
                n_BitCount += 1
                
                if n_Data == 0 :
                    break
                
        except :
            print format_exc()
            return None
        
        return n_BitCount
    
    
    #    n_BinaryString                           [IN]                        Convert Target Binary String ( ex: \x03 )
    #    n_Size                                   [IN]                        Convert Target Binary Size
    #                                             [OUT]                       Converted Unsigned Integer 
    @classmethod
    def fnConvertBinaryString2Int(cls, n_BinaryString, n_Size):
        
        try :
            
            s_Pattern = ""
            if   n_Size == 1    :   s_Pattern = 'B'     # B - Unsigned Char     , b - Signed Char
            elif n_Size == 2    :   s_Pattern = 'H'     # H - Unsigned Short    , h - Signed Short
            elif n_Size == 4    :   s_Pattern = 'I'     # I - Unsigned Integer  , i - Signed Integer
            else                :   print "\t\t\t[-] Error - ConvertBinaryString2Int ( Do not Support Convert Size : 0x%08X )" % n_Size; return None
            
            n_unpack = unpack( s_Pattern, n_BinaryString )  # type( n_unpack ) -> tuple
            return n_unpack[0]
        
        except :
            print format_exc()
            return None
            




 
    
