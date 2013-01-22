# -*- coding:utf-8 -*-



import re
from traceback import format_exc
from string import letters, digits
from struct import unpack
from collections import namedtuple

# for Class Decode
from struct import pack
from zlib import decompress
from cStringIO import StringIO
from binascii import unhexlify


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
            

    #    s_Buffer                                        [IN]                        Target String Buffer 
    #    s_Filtered_Buffer                               [OUT]                       Filtered String Buffer
    def fnTakeOffHTMLAscii(self, s_Buffer):
        
        try :
            
            s_Filtered_Buffer = ""
            
            while True :
                s_FindHTML = re.search( '&#[0-9]{2,8};', s_Buffer )
                if s_FindHTML == None :
                    break               # No Position
                
                s_Filtered_Buffer += s_Buffer[:s_FindHTML.start()]
                s_Filtered_Buffer += chr( int(s_FindHTML.group()[2:-1]))
                s_Buffer = s_Buffer[ s_FindHTML.end() ]
                if len( s_Buffer ) <= s_FindHTML.end() :
                    break               
        
        except :
            print format_exc()
            return None
        
        return s_Filtered_Buffer
    
    
    #    s_Buffer                                        [IN]                        Target String Buffer 
    #    s_Filtered_Buffer                               [OUT]                       Filtered String Buffer
    def fnTakeOffAscii(self, s_Buffer):
        
        try :
            
            s_Filtered_Buffer = ""
            
            while True :
                s_FindAscii = re.search( '#[0-9]{2,3}', s_Buffer )
                if s_FindAscii == None :
                    break               # No Position
                
                if int( s_FindAscii.group()[1:len( s_FindAscii.group() )]) < 256 :
                    s_Buffer = re.sub( s_FindAscii.group(), chr( int( s_FindAscii.group()[1:len( s_FindAscii.group() )] )), s_Buffer)
            
            s_Filtered_Buffer = s_Buffer   
        
        except :
            print format_exc()
            return None
        
        return s_Filtered_Buffer


    #    s_Buffer                                    [IN]                            Target Buffer String
    #    n_Length                                    [IN]                            
    #    s_Stream                                    [OUT]                           Extracted Buffer Stream    
    def fnTakeOffNewLineChar(self, s_Buffer, n_Length):
        
        try :
            
            s_Stream = s_Buffer
            
            n_garbage = s_Buffer.__len__() - n_Length
            if n_garbage == 0 or n_garbage < 0 :
                return s_Stream
            elif n_garbage > 0 :
                for n_Index in range( s_Buffer.__len__() - 1 ) :
                    if n_garbage == 0 :
                        break
                    if s_Buffer[n_Index] == "\r" or s_Buffer[n_Index] == "\n" :
                        s_Stream = s_Buffer[1:]
                        n_garbage -= 1
                    else :
                        break
                
                for n_Index in range( s_Buffer.__len__() - 1, 0, -1) :
                    if n_garbage == 0 :
                        break
                    if s_Buffer[n_Index] == "\r" or s_Buffer[n_Index] == "\n" :
                        s_Stream = s_Buffer[1:]
                        n_garbage -= 1
                    else :
                        break
            
        except :
            print format_exc()
            return None
        
        return s_Stream

 
      
class CDecode():
    def fnFlateDecode(self, stream):
        
        try :
            
            if ord( stream[0] ) == 13 and ord( stream[1] ) == 10 :      stream = stream[2:]
            if ord( stream[-2] ) == 13 and ord( stream[-1] ) == 10 :    stream = stream[:-2]
            return decompress( stream )
        
        except :
            
            
            Buffer = CBuffer()
            if ord( stream[0] ) == 13 and ord( stream[1] ) == 10 :      stream = stream[2:]
            if ord( stream[-2] ) == 13 and ord( stream[-1] ) == 10 :    stream = stream[:-2]
            s_Stream =  Buffer.fnTakeOffNewLineChar(stream, 0)
            return decompress( s_Stream )

    
    def fnASCII85Decode(self, stream):
        # http://pyew.googlecode.com/hg-history/e984a67f8cf1a564b97187171c237da98ce5b255/plugins/pdf.py
        
        try : 
            retval = ""
            group = []
            x = 0
            hitEod = False
            # remove all whitespace from data
            stream = [y for y in stream if not (y in ' \n\r\t')]
            while not hitEod:
                c = stream[x]
                if len(retval) == 0 and c == "<" and stream[x+1] == "~":
                    x += 2
                    continue
                # elif c.isspace():
                #    x += 1
                #    continue
                elif c == 'z':
                    assert len(group) == 0
                    retval += '\x00\x00\x00\x00'
                    continue
                elif c == "~" and stream[x+1] == ">":
                    if len(group) != 0:
                        # cannot have a final group of just 1 char
                        assert len(group) > 1
                        cnt = len(group) - 1
                        group += [ 85, 85, 85 ]
                        hitEod = cnt
                    else:
                        break
                else:
                    c = ord(c) - 33
                    assert c >= 0 and c < 85
                    group += [ c ]
                if len(group) >= 5:
                    b = group[0] * (85**4) + \
                        group[1] * (85**3) + \
                        group[2] * (85**2) + \
                        group[3] * 85 + \
                        group[4]
                    assert b < (2**32 - 1)
                    c4 = chr((b >> 0) % 256)
                    c3 = chr((b >> 8) % 256)
                    c2 = chr((b >> 16) % 256)
                    c1 = chr(b >> 24)
                    retval += (c1 + c2 + c3 + c4)
                    if hitEod:
                        retval = retval[:-4+hitEod]
                    group = []
                x += 1
            return retval
        
        except : 
            n = b = 0
            out = ''
            for c in stream:
                if '!' <= c and c <= 'u':
                    n += 1
                    b = b*85+(ord(c)-33)
                    if n == 5:
                        out += pack('>L',b)
                        n = b = 0
                    elif c == 'z':
                        assert n == 0
                        out += '\0\0\0\0'
                    elif c == '~':
                        if n:
                            for _ in range(5-n):
                                b = b*85+84
                            out += pack('>L',b)[:n-1]
                            break
            return out


    def fnASCIIHexDecode(self, stream):
        return unhexlify(''.join([c for c in stream if c not in ' \t\n\r']).rstrip('>'))
    
    
    def fnRunLengthDecode(self, stream):
        f = StringIO(stream)
        decompressed = ''
        runLength = ord(f.read(1))
        while runLength:
            if runLength < 128:
                decompressed += f.read(runLength + 1)
            if runLength > 128:
                decompressed += f.read(1) * (257 - runLength)
            if runLength == 128:
                break
            runLength = ord(f.read(1))
    #    return sub(r'(\d+)(\D)', lambda m: m.group(2) * int(m.group(1)), data)
        return decompressed
    
    def fnLZWDecode(self, stream):
        return ''.join(LZWDecoder(StringIO(stream)).run())



class LZWDecoder(object):
    def __init__(self, fp):
        self.fp = fp
        self.buff = 0
        self.bpos = 8
        self.nbits = 9
        self.table = None
        self.prevbuf = None
        return

    def readbits(self, bits):
        v = 0
        while 1:
            # the number of remaining bits we can get from the current buffer.
            r = 8-self.bpos
            if bits <= r:
                # |-----8-bits-----|
                # |-bpos-|-bits-|  |
                # |      |----r----|
                v = (v<<bits) | ((self.buff>>(r-bits)) & ((1<<bits)-1))
                self.bpos += bits
                break
            else:
                # |-----8-bits-----|
                # |-bpos-|---bits----...
                # |      |----r----|
                v = (v<<r) | (self.buff & ((1<<r)-1))
                bits -= r
                x = self.fp.read(1)
                if not x: raise EOFError
                self.buff = ord(x)
                self.bpos = 0
        return v

    def feed(self, code):
        x = ''
        if code == 256:
            self.table = [ chr(c) for c in xrange(256) ] # 0-255
            self.table.append(None) # 256
            self.table.append(None) # 257
            self.prevbuf = ''
            self.nbits = 9
        elif code == 257:
            pass
        elif not self.prevbuf:
            x = self.prevbuf = self.table[code]
        else:
            if code < len(self.table):
                x = self.table[code]
                self.table.append(self.prevbuf+x[0])
            else:
                self.table.append(self.prevbuf+self.prevbuf[0])
                x = self.table[code]
            l = len(self.table)
            if l == 511:
                self.nbits = 10
            elif l == 1023:
                self.nbits = 11
            elif l == 2047:
                self.nbits = 12
            self.prevbuf = x
        return x

    @classmethod
    def run(self):
        while 1:
            try:
                code = self.readbits(self.nbits)
            except EOFError:
                break
            x = self.feed(code)
            yield x
        return 
 
 
 

 