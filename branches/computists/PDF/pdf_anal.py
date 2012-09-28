'''
Created on 2012. 7. 13.

@author: Hong
'''

import re
import os, sys
import zlib
import binascii
import cStringIO


#################################
##### Ver.1.01 Updated data #####

class pdf_read() :
    def __init__(self):
        self.obj_data = []
        self.trailer_att = ""
        self.pdf_ver = ""
        self.obj_continue = 0
        self.stream_continue = 0
        self.trailer_continue = 0
        self.obj_count = -1
        self.fff = open("debug.txt",'wb') ## for debugging
        
        self.obj_tree = []

    def file_read(self,f):
        str_data = ""
        pFile = -1
        
        while(True) :
            file_data = f.read(1)
            pFile += 1
            if file_data == "":
                break    
                                  
            elif self.obj_continue == 1 and file_data == '#' and self.stream_continue == 0: # for Hex
            #elif self.stream_continue == 0 :
                if file_data == '#': # for Hex
                    hex_data_1 = f.read(1)         
                    hex_data_2 = f.read(1)
                    pFile += 2
                    #if re.match('[a-z]|[0-9]',hex_data_1) and re.match('[a-z]|[0-9]',hex_data_2) : 
                    if re.match('[0-9]',hex_data_1) and re.match('[a-z]|[0-9]',hex_data_2) : 
                        #print(hex_data_1,hex_data_2,self.stream_continue,self.obj_continue)
                        num_1 = int(hex_data_1)
                        if ord(hex_data_2) > 96 :
                            num_2 = ord(hex_data_2)-87
                        else :
                            num_2 = int(hex_data_2)
                        hex_data = num_1*6 + num_1*10 + num_2
                        #print(chr(hex_data))
                        str_data += chr(hex_data)
                        continue
                    else :
                        str_data = str_data + file_data + hex_data_1 + hex_data_2
            
            else :
                file_data = ord(file_data)
                            
                #if file_data == 37 or file_data == 60 or file_data == 62 or file_data == 10 or file_data == 13 :## "%", "<", ">"
                if file_data in [37, 60, 62, 10, 13] :## "%", "<", ">"
                    self.fff.write(str(str_data)+"\t"+str(self.obj_continue)+"\t"+str(self.stream_continue)+"\t"+str(self.trailer_continue)+"\n") # debugging 
                    if len(str_data) == 0 and self.stream_continue == 0:# EOF
                        str_data = ""
                        continue
                    
                    else :

                            
                        #if self.stream_continue == 1 and not file_data == 10 and not file_data == 13:
                            #str_data = str_data + chr(file_data)
                            
                        self.stack_data(str_data,pFile)
                        str_data = ""
                        
                else :
                    if self.stream_continue == 1 :## removing garbage data in stream
                        if self.obj_data[self.obj_count][2][1] == 1:
                            while(True):
                                if file_data == 10 or file_data == 13:
                                    file_data = f.read(1)
                                    file_data = ord(file_data)
                                    pFile += 1
                                else :
                                    self.obj_data[self.obj_count][2][0] = pFile
                                    self.obj_data[self.obj_count][2][1] = 0
                                    break
                        else :
                            None
                    str_data += chr(file_data)
                    
        self.fff.close()
    
                  
    def stack_data(self,data,pFile):        
        if re.match('^PDF-[0-9]\.[0-9]', data):
            self.pdf_ver = data
        else :
            if re.match('[0-9]+ [0-9] obj', data): ## Get a Object Number
                num_data = data.split() ## extract number
                self.obj_count += 1
                self.obj_continue = 1
                self.obj_data.append([num_data[0],"",""])
                return 0
            
            if 'endstream' in data:
                #if len(data) > 11: # when the number is set under 11, "endstream" string insert the last one
                    #self.obj_data[self.obj_count][2].append(data)
                    #self.obj_data[self.obj_count][2] = self.obj_data[self.obj_count][2] + data 
                self.stream_continue = 0
                self.obj_data[self.obj_count][2][1] = pFile-10
                

                return 0
            else :
                if 'stream' in data:
                    self.stream_continue = 1
                    self.obj_data[self.obj_count][2] = [pFile+1,1]
                    
                    return 0

            if 'endobj' in data:
                self.obj_continue = 0
                return 0

                
            if 'trailer' in data :
                self.trailer_continue = 1
                return 0
            
            if 'startxref' in data :
                self.trailer_continue = 0
                return 0
            
            if self.obj_continue == 1 or self.stream_continue == 1 or self.trailer_continue == 1:
                if self.stream_continue == 1:
                    #self.obj_data[self.obj_count][2].append(data)
                    None
                    #self.obj_data[self.obj_count][2] = self.obj_data[self.obj_count][2] + data
                elif self.trailer_continue == 1 :
                    #self.trailer_att.append(data)
                    self.trailer_att += data
                else :
                    #self.obj_data[self.obj_count][1].append(data)
                    self.obj_data[self.obj_count][1] += data
          
#######################################################
    
   
class find_obj():
    def __ini__(self,pdf_data):
        self.obj_data = pdf_data
        self.obj_tree = None
        
    def anal_attribute(self,data):
        pFirst = 0
        pLast = 0
        gotit = 0
        count = 0
        tree = []
        att_open = 0
        
        for i in data:
            if i == "/":
                att_open = 1
                if att_open == 1 and gotit == 1:
                    #print(data[pFirst:pLast])
                    tree.append(data[pFirst:pLast])
                    pFirst = pLast
                    #att_open = 0
                    gotit = 0
                #print(data[pFirst+1:pLast-1])
                
                if data[pFirst+1:pLast] == 'Filter' or data[pFirst+1:pLast-1] == 'Filter':
                    # '+1','-1' are used to remove garbage
                    # if they aren't there => /Filter[ 
                    gotit = 1
                    
            
            #print(i, att_open, gotit, count+1, len(data))
            #if att_open == 1 and gotit == 1 and (i == ">" or (count+1) == len(data)):
            if att_open == 1 and (i == ">" or (count+1) == len(data)):
                #print(data[pFirst:pLast+1])
                tree.append(data[pFirst:pLast+1])                    
                pFirst = pLast
                att_open = 0
                gotit = 0
                
            if i == " ":
                gotit = 1
            
            count += 1
            pLast = count
                    
        return tree
    
    def remove_chr(self,data):
        if 47 < ord(data[0]) < 58 :
            return data
        else :
            pStart = 0
            for i in data :
                if not 47 < ord(i) < 58 :
                    pStart += 1
                else :
                    return data[pStart:]
                
    
    def find_main(self,title,obj_att,stored,find_obj):
        print "[*] %s Object searching..."%find_obj

        ###########################
        ##### Making Obj Tree #####       
        att = self.anal_attribute(obj_att)
        stored = [[title,att,self.find_next_obj(att)]]
        str_att = str(att)
        if not str_att.find(find_obj) == -1 :
            self.obj_tree = stored
        else :
            if self.find_object(stored,find_obj):
                print("[-] Got %s Object!\n"%find_obj)
                return 1
            else :
                print("[-] There isn't no %s Object\n"%find_obj)
                return 0 
            
            
    def find_object(self,stored,find_obj):
        level = len(stored) - 1
        obj_count = len(stored[level][2])
        if obj_count == 0 :
            return 0
        else :
            for i in stored[level][2] :
                obj_line = self.find_obj_data_num(i)
                att = self.anal_attribute(self.obj_data[obj_line][1])
                #stored = [[i,att,self.find_next_obj(att)]]
                stored.append([i,att,self.find_next_obj(att)])

                str_att = str(att)
                #print("[*check*]"+str_att, str_att.find(find_obj))
                if not str_att.find(find_obj) == -1 :
                    self.obj_tree = stored
                    return 1                
                else :
                    if not self.find_object(stored,find_obj) == 1:
                        removed_data = stored.pop()
                        #print("[*Remove*]",removed_data)
                        None
                    else :
                        return 1
                        
                        
            
    def find_obj_data_num(self,number):
        
        for i in range(len(self.obj_data)) :
            if self.obj_data[i][0] == str(number): ## number is stored as str in obj_data
                return i
        return -1 


    def find_next_obj(self, att):
        ret_data = []
        for i in att :
            i = i.replace("[","[ ")
            i = i.replace("]"," ]")
        
            data = i.split()
            if "R" in data :
                count = data.index("R")
                if not "/Parent" in data[count-3]: # parent Object isn't displayed
                    ret_data.append(self.remove_chr(data[count-2]))
        return ret_data
    
    def find_filter(self,data): #data = att
        filter = []
        if not data.find('Filter') == -1 :
            #filter_data = data.split('/')
            filter_data = re.split('[/ \[\]]+',data) # This regular expression can help to find the string "Filter"
            
            pStart = filter_data.index('Filter')
            if 'Type' in data :
                pEnd = filter_data.index('Type')
            elif 'Length' in data :
                pEnd = filter_data.index('Length')
            else :
                pEnd = len(filter_data) - 1
            
            for i in range(pStart+1,pEnd):
                filter.append(filter_data[i])
            
            return filter
            
        else :
            print("This attribute has no filter.\n")
            return 0
        
    def decode_all(self,obj_data):
        self.obj_data = obj_data
        value = 0
        for i in range(len(self.obj_data)):
            self.obj_data[i].append(self.decoding_string(self.obj_data[i][1], self.obj_data[i][2]))
            #print(self.obj_data[i][3])
            value = 1
            
            
            
        self.obj_data = obj_data    
        return value
    
    def decode_obj(self,num,file_data):
        print(num)
        print(type(num))
        #obj_num = self.find_obj_data_num(num)
        obj_num = num
        print("obj num : %s"%obj_num)
        filter = self.obj_data[obj_num][1]
        pointer = self.obj_data[obj_num][2]
        decode_data = self.decoding_string(filter, pointer, file_data)
        
        return decode_data
        
    def decoding_string(self,filter,pointer,file_whole_data):
        #print(filter)
        
        pStart = 0
        pEnd = 0
        
        check = str(type(pointer))
        if not check.find('list') == -1:
            if pointer[0] == 0:
                "The stream Start pointer set Wrong\n"
                return 0
            elif pointer[1] == 0:
                "The stream End pointer set Wrong\n"
                "This pointer will be changed to end of file\n"
                pStart = pointer[0]
                pEnd = None
            else :
                pStart = pointer[0]
                pEnd = pointer[1]
        else :
            print("The stream pointer list has some problem\n")
            return 0
        
        filter = re.split('[/ \[\]]+',filter)
        
        data = file_whole_data[pStart:pEnd]
        print(pStart,pEnd)
        
        for i in filter :
            if "FlateDecode" in i:
                try:
                    data = FlateDecode(data)
                    print("[*] FlateDecoding Success!\n")
                except:
                    data = "decoding failed\n"
                    print("[*] FlateDecoding failed!\n")
            if "ASCII85Decode" in i:
                try:
                    data = ASCII85Decode(data)
                except:
                    data = ASCII85Decode2(data)
                print("[*] ASCII85Decoding Success!\n")
            if "ASCIIHexDecode" in i:
                data = ASCIIHexDecode(data)
                print("[*] ASCIIHexDecoding Success!\n")
            if "LZWDecode" in i:
                data = LZWDecode(data)
                print("[*] LZWDecoding Success!\n")
            if "RunLengthDecode" in i:
                data = RunLengthDecode(data)
                print("[*] RunLengthDecoding Success!\n")
                
        return data
        
        
def isJavascript(content):
    JSStrings = ['var ',';',')','(','function ','=','{','}','if ','else','return','while ','for ',',','eval']
    keyStrings = [';','(',')']
    stringsFound = []
    limit = 15
    results = 0
    
    for char in content:
        if ord(char) >= 128:
            return False

    for string in JSStrings:
        cont = content.count(string)
        results += cont
        if cont > 0:
            stringsFound.append(string)
        elif cont == 0 and string in keyStrings:
            return False

    if results > limit and len(stringsFound) >= 5:
        return True
    else:
        return False        

#############################################################
##### Deciding Filter from pdf-parser by Didier Stevens #####

## r/source/browse/trunk/pdfminer/pdfminer/ascii85.py
def ASCII85Decode2(data): ## there is some problem for the detail, Sometimes a last chr should be a '>' But, chr '=' existed at there 
    import struct
    n = b = 0
    out = ''
    for c in data:
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

def ASCII85Decode(data): ## from http://pyew.googlecode.com/hg-history/e984a67f8cf1a564b97187171c237da98ce5b255/plugins/pdf.py
    retval = ""
    group = []
    x = 0
    hitEod = False
    # remove all whitespace from data
    data = [y for y in data if not (y in ' \n\r\t')]
    while not hitEod:
        c = data[x]
        if len(retval) == 0 and c == "<" and data[x+1] == "~":
            x += 2
            continue
        #elif c.isspace():
        #    x += 1
        #    continue
        elif c == 'z':
            assert len(group) == 0
            retval += '\x00\x00\x00\x00'
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
            retval += (c1 + c2 + c3 + c4)
            if hitEod:
                retval = retval[:-4+hitEod]
            group = []
        x += 1
    return retval

def ASCIIHexDecode(data):
    return binascii.unhexlify(''.join([c for c in data if c not in ' \t\n\r']).rstrip('>'))

def FlateDecode(data):
    #return zlib.decompress(data)
    return zlib.decompress(data.strip('\r').strip('\n'))

def RunLengthDecode(data):
    f = cStringIO.StringIO(data)
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

#### LZW code sourced from pdfminer
# Copyright (c) 2004-2009 Yusuke Shinyama <yusuke at cs dot nyu dot edu>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
# documentation files (the "Software"), to deal in the Software without restriction, including without limitation 
# the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, 
# and to permit persons to whom the Software is furnished to do so, subject to the following conditions: 

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

    def run(self):
        while 1:
            try:
                code = self.readbits(self.nbits)
            except EOFError:
                break
            x = self.feed(code)
            yield x
        return

####

def LZWDecode(data):
    return ''.join(LZWDecoder(cStringIO.StringIO(data)).run())
            