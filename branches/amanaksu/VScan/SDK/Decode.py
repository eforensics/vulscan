# External Import Module
from traceback import format_exc
from zlib import decompress, error
from cStringIO import StringIO
import binascii

# Internal Import Module
from CommonSDK import CFile


class CDecode():
    @classmethod
    def fnHexray(cls, s_buffer):
        # [IN]    s_buffer
        # [OUT]   s_outbuffer
        #    - [SUCCESS] out buffer string
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            
        except :
            print format_exc()
            return None
        
        return s_outbuffer
    @classmethod
    def fnFlateDecode(cls, s_buffer, l_Param):
        # [IN] s_buffer
        #    - Parameter : Yes
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            s_outbuffer = decompress( s_buffer )
            print "[-] FlateDecode decompress successed!"
            
        except :
            print "[-] FlateDecode decompress failed"
            return None
        
        return s_outbuffer
    @classmethod
    def fnASCII85Decode(cls, s_buffer, l_Param):
        # Referred : http://pyew.googlecode.com/hg-history/e984a67f8cf1a564b97187171c237da98ce5b255/plugins/pdf.py
        # [IN] s_buffer
        #    - Parameter : No
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            group = []
            x = 0
            hitEod = False
            # remove all whitespace from data
            data = [y for y in s_buffer if not (y in ' \n\r\t')]
            while not hitEod:
                c = data[x]
                if len(s_outbuffer) == 0 and c == "<" and data[x+1] == "~":
                    x += 2
                    continue
                #elif c.isspace():
                #    x += 1
                #    continue
                elif c == 'z':
                    assert len(group) == 0
                    s_outbuffer += '\x00\x00\x00\x00'
                    continue
                elif c == "~" and data[x+1] == ">":
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
                    s_outbuffer += (c1 + c2 + c3 + c4)
                    if hitEod:
                        s_outbuffer = s_outbuffer[:-4+hitEod]
                    group = []
                x += 1

        except :
            print "[-] ASCII85Decode decompress failed"
            return None
        
        return s_outbuffer
    @classmethod
    def fnASCIIHexDecode(cls, s_buffer, l_Param):
        # [IN] s_buffer
        #    - Parameter : No
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            s_outbuffer = binascii.unhexlify(''.join([c for c in s_buffer if c not in ' \t\n\r']).rstrip('>'))
            
        except :
            print "[-] ASCIIHexDecode decompress failed"
            return None
        
        return s_outbuffer
    @classmethod
    def fnRunLengthDecode(cls, s_buffer, l_Param):
        # [IN] s_buffer
        #    - Parameter : No
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            
            f = StringIO(s_buffer)
            runLength = ord(f.read(1))
            while runLength:
                if runLength < 128:
                    s_outbuffer += f.read(runLength + 1)
                if runLength > 128:
                    s_outbuffer += f.read(1) * (257 - runLength)
                if runLength == 128:
                    break
                runLength = ord(f.read(1))
        #    return sub(r'(\d+)(\D)', lambda m: m.group(2) * int(m.group(1)), data)
            
        except :
            print "[-] RunLengthDecode decompress failed"
            return None
        
        return s_outbuffer
    @classmethod
    def fnCCITTFaxDecode(cls, s_buffer, l_Param):
        # [IN] s_buffer
        #    - Parameter : Yes
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            
        except :
            print "[-] CCITTFaxDecode decompress failed"
            return None
        
        return s_outbuffer
    @classmethod
    def fnDCTDecode(cls, s_buffer, l_Param):
        # [IN] s_buffer
        #    - Parameter : Yes
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            
        except :
            print "[-] DCTDecode decompress failed"
            return None
        
        return s_outbuffer
    @classmethod
    def fnJBIG2Decode(cls, s_buffer, l_Param):
        # [IN] s_buffer
        #    - Parameter : Yes
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            s_outbuffer = ""
            
        except :
            print "[-] JBIG2Decode decompress failed"
            return None
        
        return s_outbuffer
    @classmethod
    def fnLZWDecode(cls, s_buffer, l_Param):
        # [IN] s_buffer
        #    - Parameter : Yes
        # [OUT] s_outbuffer
        #    - [SUCCESS] Decoded string data
        #    - [FAILURE] ""
        #    - [ERROR] None
        try :
            
            LZWDecoder = CLZWDecoder()
            
            s_outbuffer = ""
            s_outbuffer = ''.join(LZWDecoder(StringIO(s_buffer)).run())
            
        except :
            print "[-] LZWDecode decompress failed"
            return None
        
        return s_outbuffer
    
class CLZWDecoder(object):
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




