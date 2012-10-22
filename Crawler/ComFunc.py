# -*- coding:utf-8 -*-

# import Public Module
import struct, traceback, glob, os, string
import zlib, cStringIO, binascii, re

# Related File Read/Write
# Property : ClassMethod
class BufferControl():
    
    @classmethod
    def Read(cls, pBuf, Position, Size):
        return pBuf[Position:Position+Size] 
    
    @classmethod
    def ReadWord(cls, pBuf, Position):
        return struct.unpack("<h", pBuf[Position:Position+2])[0]
    
    @classmethod
    def ReadDword(cls, pBuf, Position):
        return struct.unpack("<L", pBuf[Position:Position+4])[0]
    
    @classmethod
    def ReadSectorByBuffer(cls, pBuf, SecID, Size):
        try :
            Sec_Data = ""
            if Size == 0x200 :
                Sec_Position = ( SecID + 1 ) * Size
            elif Size == 0x40 :
                Sec_Position = SecID * Size
            else :
                return Sec_Data
        
            Sec_Data = BufferControl.Read(pBuf, Sec_Position, Size)
        except :
            print traceback.format_exc()
        
        return Sec_Data
    
    @classmethod
    def ReadShortByBuffer(cls, pBuf, SecID, Size):
        Short_Data = ""
        Short_Position = SecID * Size
        try :
            Short_Data = BufferControl.Read(pBuf, Short_Position, Size)
        except :
            print traceback.format_exc()
        
        return Short_Data    
    
    @classmethod
    def ReadSectorByPointer(cls, fp, SecID, Size):
        Sec_Data = ""
        Sec_Position = ( SecID + 1 ) * Size
        
        try :
            fp.seek( Sec_Position )
            Sec_Data = fp.read( Size )
        except :
            print traceback.format_exc()
        
        return Sec_Data
    
    @classmethod
    def ConvertBinary2List(cls, pBuf, Length):
        Buf_List = []
        index = 0
        Position = 0
        
        try : 
            while index < ( len(pBuf) / Length ) :
                if Length == 2 :
                    Buf_List.append( BufferControl.ReadWord(pBuf, Position) )
                elif Length == 4 :
                    Buf_List.append( BufferControl.ReadDword(pBuf, Position) )
                else :
                    print "[Except] Add Case in ConvertBinary2List( )"
                
                Position += Length
                index += 1
        except :
            print traceback.format_exc()
        
        return Buf_List 
    
    
    @classmethod
    def ExtractAlphaNumber(cls, pBuf):
        try :
            Pattern = string.letters + string.digits
            ExtList = []
            for OneByte in pBuf :
                if ( OneByte in Pattern ) :
                    ExtList.append( OneByte )
            
            OutBuf = ''.join( ExtList )
            return OutBuf
        except :
            print traceback.format_exc() 
            

    @classmethod
    def PrintBuffer(cls, File, pBuf, Position, Size):
        try :
            File['logbuf'] += "\n\t\t+" + "-" * 78 + "\n\t"
            
            Cnt = 0
            pReadBuf = BufferControl.Read(pBuf, Position, Size)
            for Byte in pReadBuf :
                File['logbuf'] += "\t 0x%s" % binascii.b2a_hex(Byte)
                Cnt += 1
                if Cnt == 10 :
                    File['logbuf'] += "\n\t"
                    Cnt = 0 
            
            File['logbuf'] += "\n\t\t+" + "-" * 78 + "\n"
            
        except:
            print traceback.format_exc()
    
    
    
class FileControl():
    @classmethod
    def OpenFileByBinary(cls, fname):
        try : 
            fp = open( fname, 'rb' )
        except :
            print traceback.format_exc()
            
        return fp
    
    @classmethod
    def ReadFileByBinary(cls, fname):
        try :
            fp = FileControl.OpenFileByBinary(fname)
            
            pBuf = "" 
            pBuf = fp.read()
            
            fp.close()
        except :
            print traceback.format_exc()
        
        return pBuf
    
    
    @classmethod
    def WriteFile(cls, fname, Buf):
        try : 
            with open( fname, 'wb' ) as fp :
                fp.write( Buf )
        except :
            print fname
            print traceback.format_exc()
    
    @classmethod
    def DeleteFile(cls, ext):
        try : 
            flist = glob.glob( ext )
            for f in flist :
                os.remove(f)
        except :
            print traceback.format_exc()
            
            

class DecodeControl():
    def FlateDecode(self, fname, stream):
        try :
            out = ""
            
            if ord( stream[0] ) == 13 and ord( stream[1] ) == 10 :      stream = stream[2:]
            if ord( stream[-2] ) == 13 and ord( stream[-1] ) == 10 :    stream = stream[:-2]          
            
#            out = zlib.decompress(stream)
#            if out != "" : 
#                return out 
                        
            out = zlib.decompress(stream.strip("\r").strip("\n"))
            if out != "" :
                return out
            
        except zlib.error :
            FileControl.WriteFile("%s_Zlib_0x%02X_0x%02X_%d.dump" % (fname, ord(stream[0]), ord(stream[1]), len(stream)), stream)
            print "[ZlibError] FlateDecode( 0x%02X 0x%02X - Length : %d )\n" % (ord(stream[0]), ord(stream[1]), len(stream))   
            print traceback.format_exc()
            return out
        
        except :
            FileControl.WriteFile("%s_Except_0x%02X_0x%02X_%d.dump" % (fname, ord(stream[0]), ord(stream[1]), len(stream)), stream)
            print "[Unknown]  Error FlateDecode( )\n\t"
            print traceback.format_exc()
            return out
        
        
    def ASCII85Decode2(self, fname, stream):
        n = b = 0
        out = ''
        for c in stream:
            if '!' <= c and c <= 'u':
                n += 1
                b = b*85+(ord(c)-33)
                if n == 5:
                    out += struct.pack('>L',b)
                    n = b = 0
                elif c == 'z':
                    assert n == 0
                    out += '\0\0\0\0'
                elif c == '~':
                    if n:
                        for _ in range(5-n):
                            b = b*85+84
                        out += struct.pack('>L',b)[:n-1]
                        break
        return out

    
    def ASCII85Decode(self, fname, stream):
    # http://pyew.googlecode.com/hg-history/e984a67f8cf1a564b97187171c237da98ce5b255/plugins/pdf.py
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

    
    def ASCIIHexDecode(self, fname, stream):
        return binascii.unhexlify(''.join([c for c in stream if c not in ' \t\n\r']).rstrip('>'))
    
    
    def RunLengthDecode(self, fname, stream):
        f = cStringIO.StringIO(stream)
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

    
    def unescape(self, escapedBytes):
        global urlsFound 
        unescapedBytes = ''
        for i in range(2,len(escapedBytes)-1,6):
            unescapedBytes += chr(int(escapedBytes[i+2]+escapedBytes[i+3],16))+chr(int(escapedBytes[i]+escapedBytes[i+1],16))
        urls = re.findall('https?://.*$', unescapedBytes, re.DOTALL)
        for url in urls:
            if url not in urlsFound:
                urlsFound.append(url)
        return unescapedBytes
    
    
    def getVarContent(self, js, bData):
        clearBytes = ''
        bData = bData.replace('\n','')
        bData = bData.replace('\r','')
        bData = bData.replace('\t','')
        bData = bData.replace(' ','')
        parts = bData.split('+')
        for part in parts:
            if re.match('["\'].*?["\']', part, re.DOTALL):
                clearBytes += part[1:-1]
            else:
                varContent = re.findall(part + '\s*?=\s*?(.*?)[,;]', js, re.DOTALL)
                if varContent != []:
                    clearBytes += self.getVarContent(js, varContent[0])
        return clearBytes
    
    
    def LZWDecode(self, fname, stream):
        return ''.join(LZWDecoder(cStringIO.StringIO(stream)).run())



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











    
    


