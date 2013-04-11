# External Import
import re, binascii
from traceback import format_exc


# Internal Import
try :
    from CommonSDK import CFile, CBuffer, CUtil
    from Decode import CDecode
except :
    print "[-] Error - Internal Import "
    exit(-1)


class CPDF():
    s_fname = ""
    l_Object = []           # Input : fnStructPDFExtractObj
    
    l_JSIndex = []          # Input : fnStructPDFExtractJSEx
    l_JSHdr = []            # Input : fnStructPDFExtractJSEx
    l_JSBody = []           # Input : fnStructPDFExtractJSEx
    
    l_JSIndex_Sus = []      # Input : fnStructPDFExtractJSEx
    l_JSHdr_Sus = []        # Input : fnStructPDFExtractJSEx
    l_JSBody_Sus = []       # Input : fnStructPDFExtractJSEx
    
    @classmethod
    def fnIsPDF(cls, s_buffer):
        # [IN]    s_buffer
        # [OUT]
        #    - [SUCCESS] "PDF"
        #    - [FAILURE/ERROR] None 
        try :
            
            if bool( re.match("^%PDF-[0-9]\.[0-9]", s_buffer) ) :
                return "PDF"
            else :
                return None
            
        except :
            print format_exc()
            return None
    @classmethod
    def fnScanPDF(cls, PDF, s_buffer):
        # [IN]    s_buffer
        # [OUT]   
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        try :
            b_Ret = None
            
            # Step 1. Extract Object By Regular Expression in PDF
            StructPDF = CStructPDF()
            b_Ret = StructPDF.fnStructPDFExtractObj(PDF, s_buffer)
            if b_Ret == False :
                print "[-] Failure - fnStructPDFExtractObj()"
                return False
            elif b_Ret == None :
                print "[-] Error - fnStructPDFExtractObj()"
                return False
            
            # Step 2. Extract JS in Object
            b_Ret = StructPDF.fnStructPDFExtractJSEx(PDF, PDF.l_Object)
            if b_Ret == False :
                print "[-] Failure - fnStructPDFExtractJSEx()"
                return False
            if b_Ret == None :
                print "[-] Error - fnStructPDFExtractJS()"
                return False
            
        except :
            print format_exc()
            return None
        
        return True

class CStructPDF():
    def fnStructPDFExtractObj(self, PDF, s_buffer):
        # [IN]    s_buffer
        # [OUT]   l_Object
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        l_Object = []
        try :
            
            l_Object = re.findall( REGULAR_EXPRESSION_OBJECT, s_buffer, re.DOTALL )
            if l_Object == [] :
                return False
            else :
                PDF.l_Object = l_Object
            
        except :
            print format_exc()
            return None
        
        return True
    def fnStructPDFExtractJSEx(self, PDF, l_Object):
        # [IN]    l_Object
        # [OUT]   l_JS
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        try :
            
            UtilPDF = CUtilPDF()
            
            for t_Object in l_Object :
                s_ObjIndex = t_Object[0]
                s_ObjHdr = t_Object[1]
                s_ObjBody = t_Object[2]
                
                s_DecBody = ""
                
                # Pre-Procedure
                
                
                
                # Decompress
                if s_ObjHdr.find( "/Filter" ) != -1 :
                    for n_Index in range( DECODE_NAME_FILTER_STD.__len__() ) :
                        l_Param = []
                        if UtilPDF.fnUtilFind(s_ObjHdr, DECODE_NAME_FILTER_STD[n_Index]) or UtilPDF.fnUtilFind(s_ObjHdr, DECODE_NAME_FILTER_AKA[n_Index]) :
                            if DECODE_NAME_FILTER_PARAM[n_Index] == "Yes" :
                                l_Predict = re.findall( ".{0,100}?/Predictor\s([0-9]{1,8})", s_ObjHdr, re.DOTALL )
                                if l_Predict != [] and l_Predict[0] > 1 :
                                    l_Param = eval( "UtilPDF.fnUtilParam%s" % DECODE_NAME_FILTER_AKA[n_Index] )(s_ObjHdr)
                            
                            s_DecBody = eval( "CDecode.fn%s" % DECODE_NAME_FILTER_STD[n_Index] )( s_ObjBody, l_Param )
                            break
                
                # None-Decompress    
                elif s_ObjHdr.find( "/JS" ) != -1 :
                    s_ObjBody = s_ObjBody.replace('\(','(')
                    s_ObjBody = s_ObjBody.replace('\)',')')
                    s_ObjBody = s_ObjBody.replace('\n',' ')
                    s_ObjBody = s_ObjBody.replace('\r',' ')
                    s_ObjBody = s_ObjBody.replace('\t',' ')
                    s_ObjBody = s_ObjBody.replace('\\\\','\\')
                    s_DecBody = s_ObjBody                    
                    
                else :
                    s_DecBody = s_ObjBody
                
                # Exception
                l_Length = re.findall( ".{0,100}?/Length\s([0-9]{1,8})" , s_ObjHdr, re.DOTALL )
                    
                if s_DecBody == None :
                    print "[-] Error - %s ( %s, %d )" % (s_ObjIndex, l_Length[0], s_ObjBody.__len__())
                    continue
                elif s_DecBody == "" :
                    print "[-] Failure - %s ( %s, %d )" % (s_ObjIndex, l_Length[0], s_ObjBody.__len__())
                    continue
                
                if CUtil.fnIsJavaScript( s_DecBody ) :
                    PDF.l_JSIndex.append( s_ObjIndex )
                    PDF.l_JSHdr.append( s_ObjHdr )
                    PDF.l_JSBody.append( s_DecBody )
                else :
                    PDF.l_JSIndex_Sus.append( s_ObjIndex )
                    PDF.l_JSHdr_Sus.append( s_ObjHdr )
                    PDF.l_JSBody_Sus.append( s_DecBody )
                
        except :
            print format_exc()
            return None
        
        return True


REGULAR_EXPRESSION_OBJECT = "([0-9]+ [0-9] obj)(\s{0,}<<.{0,100},*?>>)[\s%]*?stream(.*?)endstream[\s%]*?endobj"
DECODE_NAME_FILTER_STD = ["FlateDecode", "ASCII85Decode", "ASCIIHexDecode", "RunLengthDecode", "LZWDecode", "CCITTFaxDecode", "DCTDecode"]
DECODE_NAME_FILTER_AKA = ["Fl", "A85", "AHx", "RL", "LZW", "CCF", "DCT"]
DECODE_NAME_FILTER_PARAM = ["Yes", "No", "No", "No", "Yes", "Yes", "Yes"]
NEW_LINE = "\r\n"
                            
class CUtilPDF():
    def fnUtilFind(self, s_Buffer, s_Str):
        # [IN] s_Buffer
        # [IN] s_Str
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE/ERROR] False
        try :
            
            if s_Buffer.find( s_Str ) != -1 :
                return True
            else :
                return False
            
        except :
            print format_exc()
            return False
    def fnUtilParamFl(self, s_Buffer):
        # [IN]     s_Buffer
        # [OUT]    l_Param
        #    - [SUCCESS] Parameter List
        #    - [FAILURE] []
        #    - [ERROR] None
        try :
            
            l_Param = []
            l_Opt_Param = ["Colors", "BitsPerComponent", "Columns"]
            for s_Opt_Param in l_Opt_Param : 
                s_Opt_Pattern = ".{0,100}?/%s\s([0-9]{1,8})" % s_Opt_Param
                l_Param.append( re.findall( s_Opt_Pattern, s_Buffer, re.DOTALL )[0] )
            
        except :
            print format_exc()
            return None
        
        return l_Param
    def fnUtilParamLZW(self, s_Buffer):
        # [IN]     s_Buffer
        # [OUT]    l_Param
        #    - [SUCCESS] Parameter List
        #    - [FAILURE] []
        #    - [ERROR] None
        try :
            
            l_Param = []
            l_Opt_Param = ["Colors", "BitsPerComponent", "Columns", "EarlyChange"]
            for s_Opt_Param in l_Opt_Param : 
                s_Opt_Pattern = ".{0,100}?/%s\s([0-9]{1,8})" % s_Opt_Param
                l_Param.append( re.findall( s_Opt_Pattern, s_Buffer, re.DOTALL )[0] )                
            
        except :
            print format_exc()
            return None
        
        return l_Param
    def fnUtilParamCCF(self, s_Buffer):
        # [IN]     s_Buffer
        # [OUT]    l_Param
        #    - [SUCCESS] Parameter List
        #    - [FAILURE] []
        #    - [ERROR] None
        try :
            
            l_Param = []
            
            
        except :
            print format_exc()
            return None
        
        return l_Param
    def fnUtilParamDCT(self, s_Buffer):
        # [IN]     s_Buffer
        # [OUT]    l_Param
        #    - [SUCCESS] Parameter List
        #    - [FAILURE] []
        #    - [ERROR] None
        try :
            
            l_Param = []
            
            
        except :
            print format_exc()
            return None
        
        return l_Param

class CPrintPDF():
    pass




