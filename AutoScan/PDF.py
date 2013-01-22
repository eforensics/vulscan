

import re
from traceback import format_exc


from PE import CPE
from Definition import *
from Common import CFile, CDecode



class CPDF():
    
    #    s_pBuf                            [IN]                        File Full Buffer 
    #    BOOL                              [OUT]                       True / False
    @classmethod
    def fnCheck(cls, s_pBuf):
        try : 
            if bool( re.match("^%PDF-[0-9]\.[0-9]", s_pBuf) ) :
                return "PDF"
            else :
                return None      
        except :
            print format_exc()
            return None
    
    
    #    s_fname                            [IN]                         File Full Name
    #    s_pBuf                             [IN]                         File Full Buffer
    #    s_Format                           [IN]                         File Format
    #    BOOL Type                          [OUT]                        True / False
    @classmethod
    def fnScan(cls, s_fname, s_pBuf, s_Format) :
        
        print "\t[+] Scan PDF - %s" % s_fname
        
        try :
            
            dl_Suspicious = []
            
# Step 1. Extract Object by Regular Expression in PDF
            StructPDF = CStructPDF()
            #    ddl_Object[x][y][z]
            #        x - Regular Expression Pattern Type
            #            0 : REGULAR_EXPRESSION_OBJECT[0]
            #            1 : REGULAR_EXPRESSION_OBJECT[1]
            #
            #        y - Object List Number
            #            0 : First Object
            #            1 : Second Object
            #            .....
            #            n : Nnd Object
            #
            #        z - Object Data List
            #            x == 0
            #                0 : Header
            #                1 : Length
            #                2 : Body
            #            x == 1
            #                0 : Header
            #                1 : Body
            ddl_Object = StructPDF.fnExtractObject(s_pBuf)
            if ddl_Object == None :
                print "\t" * 2 + "[-] Failure - ExtractObject()"
                return False            
            elif ddl_Object == [] :
                print "\t" * 2 + "[*] Maybe New Type ( Do not found Object Streams )"
                return True


# Step 2. Extract Data ( JavaScript Element )
            MappedPDF = CMappedPDF()
            #    dl_JSElement[x][y]
            #        x - Regular Expression Pattern Type
            #            0 : REGULAR_EXPRESSION_OBJECT[0]
            #            1 : REGULAR_EXPRESSION_OBJECT[1]
            #
            #        y - Object Stream No
            #            0 : First Object Stream 
            #            1 : Second Object Stream 
            #            ......
            #            n : Nnd Object Stream 
            dl_JSElement = MappedPDF.fnExtractJSElement(s_fname, ddl_Object)
            if dl_JSElement == None :
                print "\t" * 2 + "[-] Failure - ExtractJSElement()"
                return False
            

# Step 3. Comparing Suspicious String List
            ExploitPDF = CExploitPDF()
            dl_Suspicious = ExploitPDF.fnExtractSuspicious( s_fname, s_pBuf, ddl_Object, dl_JSElement)
            if dl_Suspicious == None :
                print "\t" *2 + "[-] ExtractSuspicious()"
                return False


# Step 4. Print Suspicious Log
            #    dl_Suspicious[x][y]
            #        x - Suspicious Point
            #            0 : Event
            #            1 : Action
            #            2 : JavaScript
            #            3 : Embedded PE
            #
            #        y - Suspicious List Member
            #
            #        if x == 2 :
            #            y - Suspicious Double-List for CVENo
            #        if x == 3 :
            #            y - Embedded PE Count
            if not cls.fnPrintPDF(dl_Suspicious) :
                print "\t" * 2 + "[-] Failure - PrintPDF()"
                return False
            
        except :
            print format_exc()
            return False
        
        return True  
            
    
    #    s_fname               [IN]          File Full Name
    #    dl_Suspicious         [IN]          Suspicious Data List
    #                          [OUT]         True / False
    @classmethod
    def fnPrintPDF(cls, dl_Suspicious):
        #    dl_Suspicious[x][y]
        #        x - Suspicious Type
        #            0 : Event
        #            1 : Action
        #            2 : JavaScript Number
        #            3 : Embedded PE
        #        y - Suspicious List
        #
        #        if x == 2 :
        #            y - Suspicious List for CVENo
        try :
            
            print "\n" + "\t" * 2 + "=" * 72
            for n_Index in range( g_SuspiciousList.__len__() ) :
                if dl_Suspicious[n_Index] == [] or dl_Suspicious[n_Index] == 0 :
                    continue
                elif dl_Suspicious[n_Index] == None :
                    print "\n\t" * 2 + "[*] %-16s : Error - Check plz...." % g_SuspiciousList[n_Index]
                    continue
                
                #    Event, Action
                if g_SuspiciousList[n_Index] != "JavaScript" and g_SuspiciousList[n_Index] != "Embedded" :
                    print "\t" * 2 + "[*] %-16s : %s" % (g_SuspiciousList[n_Index], dl_Suspicious[n_Index])
                
                #     Embedded
                elif g_SuspiciousList[n_Index] == "Embedded" : 
                    print "\t" * 2 + "[*] %-16s : %d" % (g_SuspiciousList[n_Index], dl_Suspicious[n_Index])
                
                #     JavaScript
                else :  
                    for n_CVENo in dl_Suspicious[n_Index] :
                        print "\t" * 2 + "[*] %-16s : %-30s ( %s )" % (g_SuspiciousList[n_Index], g_SuspiciousJS[ n_CVENo ], g_CVENo[ n_CVENo ])
                        
            print "\t" * 2 + "=" * 72
            
        except :
            print format_exc()
            return False
        
        return True 



class CStructPDF():
    
    #    s_pBuf                [IN]                PDF Full Buffer
    #    ddl_Object            [OUT]               Object 2-Type DDouble-List in PDf
    def fnExtractObject(self, s_pBuf):

        try :
            
            ddl_Object = []
            
            for s_RE_Pattern in REGULAR_EXPRESSION_OBJECT :
                ddl_Object.append( re.findall( s_RE_Pattern, s_pBuf, re.DOTALL ) )
            
        except :
            print format_exc()
            return None
        
        return ddl_Object



class CMappedPDF():
    
    #    s_fname                 [IN]                File Full Name
    #    ddl_Object              [IN]                Object DDouble-List ( 2-Type Regular Expression )
    #    l_JSElement             [OUT]               JavaScript Element ( Stream Body ) List 
    def fnExtractJSElement(self, s_fname, ddl_Object):
        
        try :
            
            dl_JavaScript = []
            
            #    ddl_Object[x][y][z]
            #        x - Regular Expression Pattern Type
            #            0 : REGULAR_EXPRESSION_OBJECT[0]
            #            1 : REGULAR_EXPRESSION_OBJECT[1]
            #
            #        y - Object List No
            #            0 : First Object
            #            1 : Second Object
            #            .....
            #            n : Nnd Object
            #
            #        z - Object 
            #            x == 0
            #                0 : Header
            #                1 : Length
            #                2 : Body
            #            x == 1
            #                0 : Header
            #                1 : Body
            
            for dl_Object in ddl_Object :
                for s_Cmp in g_JavaScript :
                    l_ExtractJS = self.fnExtractJS(s_fname, dl_Object, s_Cmp)
                    if l_ExtractJS == None :
                        print "\t" * 3 + "[-] Failure - ExtractJS( %s )" % s_Cmp
                        return None
                    if l_ExtractJS == [] :
                        continue
                    dl_JavaScript.append( l_ExtractJS )
            
        except :
            print format_exc()
            return None
        
        return dl_JavaScript



    #    s_fname                        [IN]                    File Full Name
    #    dl_Object                      [IN]                    Object Double-List
    #    s_Cmp                          [IN]                    Comparing String ( /JS, /JavaScript )
    #    l_ExtractJS                    [OUT]                   Extract JS Object Body List
    def fnExtractJS(self, s_fname, dl_Object, s_Cmp):
        
        try :
            
            l_ExtractJS = []
            
            for n_Index in range( dl_Object.__len__() ) :
                l_Object = dl_Object[n_Index]
                if l_Object[0].find( s_Cmp ) == -1 :
                    continue
                
                n_BodyIndex = l_Object.__len__() -1
                
                #    Case 1. /JS
                if s_Cmp == "/JS" :
                    s_DecompObject = l_Object[ n_BodyIndex ]
                    s_DecompObject = s_DecompObject.replace('\(','(')
                    s_DecompObject = s_DecompObject.replace('\)',')')
                    s_DecompObject = s_DecompObject.replace('\n',' ')
                    s_DecompObject = s_DecompObject.replace('\r',' ')
                    s_DecompObject = s_DecompObject.replace('\t',' ')
                    s_DecompObject = s_DecompObject.replace('\\\\','\\')
                
                
                #    Case 2. /JavaScript
                #    Case 3. /JavaScript + g_Decode
                #    Case 4. g_Decode
                else :
                    s_DecompObject = self.fnDecompObject(l_Object)
                    if s_DecompObject == None :
                        print "\t" * 3 + "[-] Failure - DecompObject()"
                        CFile.fnWriteFile("c:\\test1\\%s_ERROR_Decomp.dump" % s_fname, str(l_Object) )
                        continue
                    if s_DecompObject == "" :
                        s_DecompObject = l_Object[ n_BodyIndex ]
            
                if s_DecompObject != "" :
                    l_ExtractJS.append( s_DecompObject )
            
        except :
            print format_exc()
            return None
        
        return l_ExtractJS


    #    l_Object                                    [IN]                            Object List
    #    s_DecompObject                              [OUT]                           Decompress Object Body Data
    def fnDecompObject(self, l_Object):
        
        try :
            
            Decode = CDecode()
            
            s_DecompObject = ""
            n_BodyIndex = l_Object.__len__() - 1            
            
            for s_Decode in g_Decode :
                if l_Object[0].find( s_Decode ) != -1 :
                    if s_Decode == "/Fl" :          s_Decode = "/FlateDecode"
                    s_DecompObject = eval( "Decode.fn" + s_Decode[1:] )( l_Object[ n_BodyIndex ] )
                    break
            
        except :
            print format_exc()
            return None
        
        return s_DecompObject





class CExploitPDF():
    
    #    s_fname                    [IN]                    File Full Name
    #    s_pBuf                     [IN]                    File Full Buffer
    #    ddl_Object                 [IN]                    Object Double-List in PDF
    #    dl_JSElement               [IN]                    JavaScript Element Double-List
    #    dl_Suspicious              [OUT]                   Suspicious Double-List
    def fnExtractSuspicious( self, s_fname, s_pBuf, ddl_Object, dl_JSElement):
    
        try :
            
            dl_Suspicious = []
            
            l_Event = self.fnExploitPDFEvent(s_pBuf)
            if l_Event == None :
                print "\t" * 2 + "[-] Failure - ExploitPDFEvent()"
                return False
            dl_Suspicious.append( l_Event )
            
            l_Action = self.fnExploitPDFAction(s_pBuf)
            if l_Action == None :
                print "\t" * 2 + "[-] Failure - ExploitPDFAction()"
                return False
            dl_Suspicious.append( l_Action )
            
            l_CVENo = self.fnExploitPDFJavaScript(s_fname, dl_JSElement)
            if l_CVENo == None :
                print "\t" * 2 + "[-] Failure - ExploitPDFJavaScript()"
                return False
            dl_Suspicious.append( l_CVENo )

            n_Embedded = self.fnExploitPDFEmbedded(s_fname, ddl_Object)
            if n_Embedded == None :
                print "\t" *2 + "[-] Failure - ExploitPDFEmbedded()"
                return False
            dl_Suspicious.append( n_Embedded )
            
        except :
            print format_exc()
            return None
        
        return dl_Suspicious
        
    
    #    s_pBuf                [IN]                PDF Full Buffer
    #    l_Event               [OUT]               Event List in PDF
    def fnExploitPDFEvent(self, s_pBuf):
        
        try :
            
            l_Event = []
            for s_Event in g_SuspiciousEvents :
                if s_pBuf.find( s_Event ) != -1 :
                    if s_Event not in l_Event :
                        l_Event.append( s_Event )
            
        except :
            print format_exc()
            return None
        
        return l_Event
    

    #    s_pBuf                [IN]                PDF Full Buffer
    #    l_Action              [OUT]               Action List in PDF    
    def fnExploitPDFAction(self, s_pBuf):
        
        try :
            
            l_Action = []
            for s_Action in g_SuspiciousActions :
                if s_pBuf.find( s_Action ) != -1 :
                    if s_Action not in l_Action :
                        l_Action.append( s_Action )
            
        except :
            print format_exc()
            return None
        
        return l_Action
    
    
    #    s_fname               [IN]                File Full Name
    #    dl_JSElement          [IN]                JavaScript Stream List in Object
    #    l_JSElement           [OUT]               JavaScript Element Number in Stream
    def fnExploitPDFJavaScript(self, s_fname, dl_JSElement):
        #    dl_JSElement[x][y]
        #        x - Regular Expression Pattern Type
        #            0 : REGULAR_EXPRESSION_OBJECT[0]
        #            1 : REGULAR_EXPRESSION_OBJECT[1]
        #
        #        y - Object Stream No
        #            0 : First Object Stream 
        #            1 : Second Object Stream 
        #            ......
        #            n : Nnd Object Stream 
        
        try :
            
            dl_CVENo_Du = []
            for l_JSElement in dl_JSElement :
                for n_Index in range( l_JSElement.__len__() ) :
                    s_Stream = l_JSElement[ n_Index ]
                    if not self.fnIsJavaScript(s_Stream) :
                        continue
                    
                    l_CVENo = []
                    l_CVENo = self.fnExploitPDFJS(s_Stream)
                    if l_CVENo == None or l_CVENo == [] :
                        s_UnescapedStream = self.fnUnescaping(s_Stream)
                        if s_UnescapedStream == None or s_UnescapedStream == "" :
                            s_dumpName = "c:\\test1\\%s_ERROR_0x%08X_Unescaping.dump" % (s_fname, n_Index)
                            print "\t" * 2 + "[-] Warning - Unescaping() [ Dumpped : %s ]" % s_dumpName
                            CFile.fnWriteFile(s_dumpName, s_Stream)
                            continue
                        l_CVENo = self.fnExploitPDFJS(s_UnescapedStream)
                        if l_CVENo == None :
                            print "\t" * 2 + "[-] Failure - ExploitPDFJS()"
                            continue
                    
                    if l_CVENo != [] :
                        dl_CVENo_Du.append( l_CVENo )
                        
            l_CVENo_NoneDu = self.fnExploitCVENoUnDuplicate(dl_CVENo_Du)
            if l_CVENo_NoneDu == None or l_CVENo_NoneDu == False :
                print "\t" * 2 + "[-] Failure - ExploitCVENoUnDuplicate()"
                return None
            
        except :
            print format_exc()
            return None
        
        return l_CVENo_NoneDu


    #    dl_CVENo               [IN]               CVE No Double-List
    #    l_CVENo                [OUT]              CVE No List ( None-Duplicate )
    def fnExploitCVENoUnDuplicate(self, dl_CVENo):
        
        try :
            
            l_CVENo_NoneDulpicate = []
            
            for n_RE_Index in range( dl_CVENo.__len__() ):
                l_CVENo_Duplicate = dl_CVENo[ n_RE_Index ]
                
                for n_CVENo in l_CVENo_Duplicate :
                    if n_CVENo not in l_CVENo_NoneDulpicate :
                        l_CVENo_NoneDulpicate.append( n_CVENo )
            
        except :
            print format_exc()
            return None
        
        return l_CVENo_NoneDulpicate


    #    s_fname                                [IN]                            File Full Name
    #    ddl_Object                             [IN]                            Object Double-List
    #    n_Cnt                                  [OUT]                           Embedded PE File Count
    def fnExploitPDFEmbedded(self, s_fname, ddl_Object):
        
        try :
            #    ddl_Object[x][y][z]
            #        x - Regular Expression Pattern Type
            #            0 : REGULAR_EXPRESSION_OBJECT[0]
            #            1 : REGULAR_EXPRESSION_OBJECT[1]
            #
            #        y - Object List No
            #            0 : First Object
            #            1 : Second Object
            #            .....
            #            n : Nnd Object
            #
            #        z - Object 
            #            x == 0
            #                0 : Header
            #                1 : Length
            #                2 : Body
            #            x == 1
            #                0 : Header
            #                1 : Body
            n_Cnt = 0
            for dl_Object in ddl_Object :
                for n_Index in range( dl_Object.__len__() -1 ) :
                    l_Object = dl_Object[ n_Index ]
                    for s_BodyObject in l_Object[ l_Object.__len__() - 1 ] :
                        if CPE.Check( s_BodyObject ) == "PE" :
                            n_Cnt += 1
                            CFile.fnWriteFile("c:\\test1\\%s_0x%08X_Embedded.dump" % (n_Index, s_fname), s_BodyObject)
                            
                            
        except :
            print format_exc()
            return None
        
        return n_Cnt 





    #    s_Stream                        [IN]                        Object Stream
    #    l_CVENo                         [OUT]                       Suspicious JavaScript CVE Number List in g_SuspiciousJS
    def fnExploitPDFJS(self, s_Stream):
        
        try :
            
            l_CVENo = []
            for n_Index in range( g_SuspiciousJS.__len__() ) :
                s_JS = g_SuspiciousJS[n_Index]
                if s_Stream.find( s_JS ) != -1 :
                    if n_Index not in l_CVENo :
                        l_CVENo.append( n_Index )
        
        except :
            print format_exc()
            return None
        
        return l_CVENo


    #    s_Stream                            [IN]                        Object Stream
    #    BOOL                                [OUT]                       True / False
    def fnIsJavaScript(self, s_Stream):
        
        try :
            l_JSStrings = ['var ',';',')','(','function ','=','{','}','if ','else','return','while ','for ',',','eval']
            l_keyStrings = [';','(',')']
            l_stringsFound = []
            n_limit = 15
            n_results = 0
            
            for s_char in s_Stream:
                if ord(s_char) >= 150:
                    return False
        
            for s_string in l_JSStrings:
                n_cont = s_Stream.count(s_string)
                n_results += n_cont
                if n_cont > 0:
                    l_stringsFound.append(s_string)
                elif n_cont == 0 and s_string in l_keyStrings:
                    return False
        
            if n_results > n_limit and l_stringsFound.__len__() >= 5:
                return True
        
            else:
                return False
        
        except :
            print format_exc()
            return False
        
    
    #    s_Stream                    [IN]                    Object Stream
    #    s_UnescapedStream           [OUT]                   Unescaped Object Stream
    def fnUnescaping(self, s_Stream):
        
        try :
            
            s_UnescapedStream = ""
            
            l_EscapedVars = re.findall(REGULAR_EXPRESSION_ESCAPE, s_Stream, re.DOTALL)
            if l_EscapedVars == [] :
                print "\t" * 2 + "[-] JavaScript is. but do not find String \"escape\" "
                return s_UnescapedStream 
            
            for s_Var in l_EscapedVars :
                s_Data = s_Var[2]
                if s_Data.find("+") != -1 :
                    s_Content = self.fnGetVarContent(s_Stream, s_Data)
                else :
                    s_Content = s_Data[1:-1]
                    
                if s_Content.__len__() > 150 : 
                    return self.fnUnescape(s_Content)
                    
        except :
            print format_exc()
            return None
        
        return s_UnescapedStream 
            
    
    #    s_Stream                    [IN]                    Object Stream
    #    s_Data                      [IN]                    "+" String Data
    #    s_CleatBytes                [OUT]                   Removed "+" String Data
    def fnGetVarContent(self, s_Stream, s_Data):
        
        try :
            
            s_ClearBytes = ""
            s_Data = s_Data.replace('\n','')
            s_Data = s_Data.replace('\r','')
            s_Data = s_Data.replace('\t','')
            s_Data = s_Data.replace(' ','')
            l_parts = s_Data.split('+')
            for s_part in l_parts:
                if re.match('["\'].*?["\']', s_part, re.DOTALL):
                    s_ClearBytes += s_part[1:-1]
                else:
                    varContent = re.findall(s_part + '\s*?=\s*?(.*?)[,;]', s_Stream, re.DOTALL)
                    if varContent != []:
                        s_ClearBytes += self.getVarContent(s_Stream, varContent[0])
                        
        except :
            print format_exc()
            return None
        
        return s_ClearBytes


    #    s_EscapedBytes                        [IN]                    Object Stream
    #    s_UnescapedBytes                      [OUT]                   Unescaped Object Stream
    def fnUnescape(self, s_EscapedBytes):
        
        try :
            global urlsFound 
            
            s_UnescapedBytes = ""
            
            for n_Index in range(2,s_EscapedBytes.__len__() -1,6):
                s_UnescapedBytes += chr(int(s_EscapedBytes[n_Index+2]+s_EscapedBytes[n_Index+3],16))+chr(int(s_EscapedBytes[n_Index]+s_EscapedBytes[n_Index+1],16))
            l_urls = re.findall('https?://.*$', s_UnescapedBytes, re.DOTALL)
            for s_url in l_urls:
                if s_url not in urlsFound:
                    urlsFound.append(s_url)
                    
        except :
            print format_exc()
            return None
                
        return s_UnescapedBytes


