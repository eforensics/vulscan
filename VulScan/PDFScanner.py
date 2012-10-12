# -*- coding:utf-8 -*-

# import Public Module
import re, traceback, binascii


# import private module
from ComFunc import FileControl, DecodeControl



class OperateStream():
    def isJavaScript(self, element):
        try :
            JSStrings = ['var ',';',')','(','function ','=','{','}','if ','else','return','while ','for ',',','eval']
            keyStrings = [';','(',')']
            stringsFound = []
            limit = 15
            results = 0
            
            for char in element:
                if ord(char) >= 150:
                    return False
        
            for string in JSStrings:
                cont = element.count(string)
                results += cont
                if cont > 0:
                    stringsFound.append(string)
                elif cont == 0 and string in keyStrings:
                    return False
        
            if results > limit and len(stringsFound) >= 5:
                return True
        
            else:
                return False
        
        except TypeError :
            return False
        
        except :
            print traceback.format_exc()
            return False


    def ExtractStream(self, stream, length):
        try :
            output = stream
            garbage = len( stream ) - length
            if garbage == 0 or garbage < 0 :
                return stream
            elif garbage > 0 :
                i = 0
                for i in range( len(stream)) :
                    if garbage == 0 :
                        break
                    if stream[i] == '\r' or stream[i] == '\n' :
                        output = output[1:]
                        garbage -= 1
                    else :
                        break
        
                i = 0
                for i in range( len(stream)-1, 0, -1 ) :
                    if garbage == 0 :
                        break
                    if stream[i] == '\r' or stream[i] == '\n' :
                        output = output[:-1]
                        garbage -= 1
                    else :
                        break
                    
                return output
            
        except :
            print traceback.format_exc()


    def Unescaping(self, element):
        try :
            Decode = DecodeControl()
        
            escapedVars = re.findall('(\w*?)\s*?=\s*?(unescape\((.*?)\))', element, re.DOTALL)
            for var in escapedVars:
                bData = var[2]
                if bData.find('+') != -1:
                    varContent = Decode.getVarContent(element, bytes)
                    if len(varContent) > 150:
                        bData = Decode.unescape(varContent)
                        return bData
                else:
                    bData = bData[1:-1]
                    if len(bData) > 150:
                        bData = Decode.unescape(bData)
                        return bData      
        
        except TypeError :
            print "\n\t\t[TypeError] Unescape( )\n\t\t- %s" % var
          
        except :
            print traceback.format_exc()
        
        return


    def FilterObj(self, element):
        try :
            outElement = element            
            tmpElement = element
            # HTML ASCII
            tmpElement = self.fltHTMLAscii( tmpElement )
            if tmpElement == "" :
                tmpElement = element
            else :
                outElement = tmpElement
            
            # ASCII
            tmpElement = self.fltAscii( tmpElement )
            if tmpElement == "" :
                tmpElement = element 
            else :
                outElement = tmpElement
            
        except :
            print traceback.format_exc()
        
        return outElement   


    def fltHTMLAscii(self, pBuf):
        try :
            HTMLAscii_Buf = ""
            Buffer = pBuf
            while True :          
                SearchElement = re.search( '&#[0-9]{2,8};', Buffer )
                
                HTMLAscii_Buf += Buffer[:SearchElement.start()]
                HTMLAscii_Buf += chr( int(SearchElement.group()[2:-1]) )
                Buffer = Buffer[SearchElement.end():]
                if len(Buffer) <= SearchElement.end() :
                    break
        
        except AttributeError :
            HTMLAscii_Buf += Buffer
    
        return HTMLAscii_Buf
    
    
    def fltAscii(self, pBuf):
        try :
            Ascii_Buf = ""
            tmpBuf = pBuf
            while True :
                SearchElement = re.search( '#[0-9]{2,3}', tmpBuf )
                if int(SearchElement.group()[1:len( SearchElement.group() )]) < 256 :
                    tmpBuf = re.sub( SearchElement.group(), chr( int(SearchElement.group()[1:len( SearchElement.group() )]) ), tmpBuf )            
        
            Ascii_Buf = tmpBuf
        
        except AttributeError :
            Ascii_Buf = tmpBuf
        
        except :
            print traceback.format_exc()
        
        return Ascii_Buf



class PDFSearch():
    def SearchEvent(self, pBuf):
        try :
            EventList = []
            for event in SuspiciousEvents :
                if pBuf.find( event ) != -1 :
                    if event not in EventList :
                        EventList.append( event )
        except :
            print traceback.format_exc()
            return []
            
        return EventList
            
    
    def SearchAction(self, pBuf):
        try : 
            ActionList = []
            for action in SuspiciousActions : 
                if pBuf.find( action ) != -1 :
                    if action not in ActionList :
                        ActionList.append( action )
        except :
            print traceback.format_exc()
            return []
            
        return ActionList   
    
    
    def SearchElement(self, File, pBuf, ActionList):
        try :
            ElementList = []      
            # Javascript is 2-types. first type is /JS. Second type is /JavaScript.
            # First Type : "/JS"
            if '/JS' in ActionList :
                JSContent = re.findall('/JS\s*?\((.*?)\)[\r\n]', pBuf, re.DOTALL)
                if JSContent != [] :
                    js = JSContent[0]
                    js = js.replace('\(','(')
                    js = js.replace('\)',')')
                    js = js.replace('\n',' ')
                    js = js.replace('\r',' ')
                    js = js.replace('\t',' ')
                    js = js.replace('\\\\','\\')
                    ElementList.append(js)
            
    
            # Second Type : "/JavaScript"
            Operation = OperateStream()
            streams = re.findall('obj\s{0,10}<<(.{0,100}?/Length\s([0-9]{1,8}).*?)>>[\s%]*?stream(.*?)endstream[^=]',pBuf,re.DOTALL)
            if streams == [] :
                pBuf = Operation.FilterObj(pBuf)
                streams = re.findall('obj\s{0,10}<<(.{0,100}?/Length\s([0-9]{1,8}).*?)>>[\s%]*?stream(.*?)endstream[^=]',pBuf,re.DOTALL)
    
            
            for streamDict in streams :
                Dict = streamDict[0]  
                
                if Dict.count('obj') > 0:
                    Dict = Dict[Dict.rfind('obj'):]
                
                stream = Operation.ExtractStream( streamDict[2], int(streamDict[1]) )
                if len(stream) < 0x10 :
                    continue
                
#                Filters : /FlateDecode, /ASCIIHexDecode, /ASCII85Decode, /LZWDecode, /RunLengthDecode
                Decode = DecodeControl()
                if Dict.find('/FlateDecode') != -1 or Dict.find('/Fl') != -1:
                    streamContent = Decode.FlateDecode(File["fname"], stream)
                    if streamContent == "" :
                        ConvertData = ""
                        for offset in range( len(stream) / 2 ) : 
                            bData = stream[offset*2:offset*2+2]
                            ConvertData += binascii.a2b_hex( bData )

                        stream = ConvertData
                        streamContent = Decode.FlateDecode(File["fname"], stream)  
            
                elif Dict.find('/ASCII85Decode') != -1 :
                    try :
#                        print "ASCII85Deocde"
                        streamContent = Decode.ASCII85Decode(File["fname"], stream)
                    except : 
#                        print "ASCII85Decode2"
                        streamContent = Decode.ASCII85Decode2(File["fname"], stream)
                            
                elif Dict.find('/ASCIIHexDecode') != -1 :
#                    print "ASCIIHexDecode"
                    streamContent = Decode.ASCIIHexDecode(File["fname"], stream)
                        
                elif Dict.find('/RunLengthDecode') != -1 :
#                    print "RunLengthDecode"
                    streamContent = Decode.RunLengthDecode(File["fname"], stream)
                        
                elif Dict.find('/LZWDecode') != -1 :
#                    print "LZWDecode"
                    streamContent = Decode.LZWDecode(File["fname"], stream)
                        
                else:
#                    print "No-Deocde"
                    streamContent = stream
                           
                ElementList.append(streamContent)

#            Add Procedure
            if streams == [] :
                streams = re.findall('obj\s{0,10}<<(.{0,100}.*?)>>[\s%]*?stream(.*?)endstream[^=]',pBuf,re.DOTALL)
                
                for streamDict in streams :
                    ElementList.append( streamDict[1] )
                                
        except :
            print traceback.format_exc()
            return []
        
        return ElementList


    def SearchJS(self, pBuf):
        try :
            JSList = []
            for js in SuspiciousJS :
                if pBuf.find( js ) != -1 :
                    if js not in JSList :
                        JSList.append( js )
        
        except AttributeError :
            return JSList
        
        except :
            print traceback.format_exc()
            
        return JSList



class PDFScan():
    @classmethod
    def Check(cls, pBuf):
        try : 
            if bool( re.match("^%PDF-[0-9]\.[0-9]", pBuf) ) :
                return True
            else :
                return False      
        except :
            print traceback.format_exc()
        
        
    @classmethod
    def Scan(cls, File):
        try :
            File['logbuf'] += "\n    [+] %s...........%s" % ( File["fname"], File["format"] )
            
            pBuf = File["pBuf"]
            Search = PDFSearch()
            
            # Check Suspicious Events
            Event = Search.SearchEvent(pBuf)
            
            # Check Suspicious Actions
            Action = Search.SearchAction(pBuf)
            
            # Check Element to Analyze with Actions
            Element = Search.SearchElement(File, pBuf, Action)
            
            # Check UnEscape to Analyze Element in PDF     
            Operation = OperateStream()
            unescapedBytes = []
            for element in Element :
                if type(element) == type(None) :
                        continue
                
#                print "Scan : %s" % File["fname"]
                element = Operation.FilterObj(element)
                if Operation.isJavaScript( element ) :
                    name = File["fname"] + "_" + hex( len(element) )
                    FileControl.WriteFile(name, element)
                    
                    UnEscapeData = Operation.Unescaping( element )
                    if UnEscapeData not in unescapedBytes :
                        unescapedBytes.append( UnEscapeData )
            
            ElementData = ""
            for element in Element :
                JSList = Search.SearchJS(element)
                if JSList == [] :
                    continue

                for JS in JSList :
                    ElementData += "[%s : %s]\n" % (JS, CVENo[JS])
                ElementData += element
            
            
            if Event == [] and Action == [] and ElementData == "" :
                File['logbuf'] += "\t[ Done ]"
            else :
                File['logbuf'] += "\t[ Find ]"
            
            
            # Logging Suspicious Data 
            if Event != [] :
                File['logbuf'] += "\n            Event  : "
                for event in Event :
                    File['logbuf'] += event + " "
                
                
            if Action != [] :
                File['logbuf'] += "\n            Action : "
                for action in Action :
                    File['logbuf'] += action + " "
                
                
            if ElementData != "" :
                FileControl.WriteFile("%s_Element.dat" % File["fname"], ElementData)    
                File['logbuf'] += "\n        [-] %s\t<%s_Element.dat>" % (File["fname"],  File["fname"]) 
            
        except :
            File['logbuf'] += "\t[ Error ]"
            print traceback.format_exc()
            return False
        
        return True


SuspiciousEvents = ['/OpenAction', '/AA']
SuspiciousActions = ['/JS', '/JavaScript', '/Launch', '/SubmitForm', '/ImportData']
SuspiciousJS = ['mailto', 
                'Collab.collectEmailInfo', 
                'util.printf', 
                'getAnnots', 
                'getIcon', 
                'spell.customDictionaryOpen', 
                'media.newPlayer']

CVENo = {"mailto"                       :   "CVE-2007-5020", 
         "Collab.collectEmailInfo"      :   "CVE-2007-5659", 
         "util.printf"                  :   "CVE-2008-2992", 
         "getAnnots"                    :   "CVE-2009-1492", 
         "getIcon"                      :   "CVE-2009-0927", 
         "spell.customDictionaryOpen"   :   "CVE-2009-1493", 
         "media.newPlayer"              :   "CVE-2009-4324"}








