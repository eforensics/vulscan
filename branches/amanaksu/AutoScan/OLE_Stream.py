
import os
from struct import unpack
from traceback import format_exc
from collections import namedtuple


from Definition import *
from Common import CBuffer, CFile
from OLE_SPM import *

class CMappedStream():
    
    #    s_fname                                          [IN]                        File Full Name
    #    s_WordDoc                                        [IN]                        WordDocument Stream Buffer
    #    s_Table                                          [IN]                        Table Stream Buffer ( 0Table or 1Table )
    #    n_Offset                                         [IN]                        Member Offset in Table Stream
    #    n_Size                                           [IN]                        Member Size in Table Stream
    #    BOOL Type                                        [OUT]                       True / False
    def fnMappedPlcBteChpx(self, s_fname, s_WordDoc, s_Table, n_Offset, n_Size):
        
        try :
            MappedBasic = CMappedBasic()
            ParseStream = CParseStream()
            
# Step 1. Parse Plc List ( Basic )
            #    l_FC        : PlcFkpChpx.aFC        - Data     Offset ( Array )
            #    l_PnFkpChpx : PlcFkpChpx.PnFkpChpx  - Property Offset ( Array ) 
            l_FC, l_PnFkpChpx = MappedBasic.fnParsePlc( s_Table, n_Offset, n_Size )
            if l_FC == None or l_PnFkpChpx == None : 
                print "\t\t\t\t[-] Failure - ParsePlc()"
                return False
            if l_FC.__len__() != l_PnFkpChpx.__len__() + 1 :
                print "\t\t\t\t[-] Error ( FC : 0x%08X, PnFkpChpx : 0x%08X )" % (l_FC.__len__(), l_PnFkpChpx.__len__())
                return False            


# Step 2. Dump PlcBteChpx
            #    l_FC_Text    : Dumpped Text Data
            #    l_ChpxFkp    : Dumpped Property
            l_FC_Text, l_ChpxFkp = ParseStream.fnParsePlcBteChpx(s_WordDoc, l_FC, l_PnFkpChpx)
            if (l_FC_Text == None or l_ChpxFkp == None) or (l_FC_Text == [] or l_ChpxFkp == []) :
                print "\t" * 4 + "[-] Failure - ParsePlcBteChpx()"
                print "\t" * 4 + "PlcBteChpx.aFC\n" + "\t" * 4 + "%s" % l_FC
                print "\t" * 4 + "PlcBteChpx.PnFkpChpx\n" + "\t" * 4 + "%s" % l_PnFkpChpx
                return False
            

# Step 3. Parse CP List ( Character Property ) with CRUN
            #    dl_Rgfc        : Double-List Rgfc per Dump
            #    dl_Rgb         : Double-List Rgb per Dump
            dl_Rgfc, dl_Rgb = MappedBasic.fnParseCPwithCRUN(l_ChpxFkp)
            if (dl_Rgfc == None or dl_Rgb == None) or (dl_Rgfc == [] or dl_Rgb == []) :
                print "\t" * 4 + "[-] Failure - ParseCP()"
                return False


## Step 4. Dump Chpx
            #    dl_Rgfc_Text    : Double-List Rgfc Text per Dump
            #    dl_Chpx         : Double-List Chpx per Dump
            dl_Rgfc_Text, dl_Chpx = ParseStream.fnParseChpx(s_WordDoc, l_ChpxFkp, dl_Rgfc, dl_Rgb)
            if (dl_Rgfc_Text == None or dl_Chpx == None) or (dl_Rgfc_Text == [] or dl_Chpx == []) :
                print "\t" * 4 + "[-] Failure - ParseChpx()"
                return False


# Step 4. Parse Chpx.GrpPrl
            B_Ret = ParseStream.fnParseGrpPrl(dl_Chpx)
            if B_Ret != True :
                print "\t" * 4 + "[-] Failure - ParseGrpPrl()"
                return False
        
        
        except :
            print format_exc()
            return False
        
        return True
        
        

class CParseStream():

    def fnParsePlcBteChpx(self, s_WordDoc, l_FC, l_PnFkpChpx):
        
        try :
            ParseStream = CParseStream()
            
            # Dump aFC 
            l_FC_Text = []
            for n_Index in range( l_FC.__len__() ) :
                if n_Index + 1 == l_FC.__len__() :
                    break
                n_Offset = l_FC[n_Index]
                n_Size = l_FC[n_Index + 1] - l_FC[n_Index]
                s_FC = s_WordDoc[ n_Offset:n_Offset + n_Size ]
                l_FC_Text.append( s_FC )
                            
            
            # Dump ChpxFkp
            l_StartChpxFkp = ParseStream.fnParsePnFkpChpx( l_PnFkpChpx )
            if l_StartChpxFkp == None or l_StartChpxFkp == [] :
                print "\t\t\t\t[-] Failure - ParsePnFkpChpx()"
                return None, None
            if s_WordDoc.__len__() < l_StartChpxFkp[ l_StartChpxFkp.__len__() - 1 ] + 0x1FF :
                print "\t\t\t\t[-] Error - Out of Size in WordDocument( WordDoc Size : 0x%08X, End Offset : 0x%08X )" % ( s_WordDoc.__len__(), l_StartChpxFkp[ l_StartChpxFkp.__len__() - 1 ] + 0x1FF )
                return None, None
            
            l_ChpxFkp = []
            for n_Offset in l_StartChpxFkp :
                s_ChpxFkp = s_WordDoc[n_Offset:n_Offset + 0x200]
                n_cron = CBuffer.fnConvertBinaryString2Int(s_ChpxFkp[-1], 1)
                if n_cron <= 0x0 or n_cron > 0x65 :
                    print "\t\t\t\t[-] Error - Read ChpxFkp.crun ( 0x%08X in WordDocument )" % n_Offset
                    CFile.fnWriteFile("%s_0x08X.dump" % (os.path.abspath(os.curdir), n_Offset) , s_ChpxFkp)
                    return None, None
                l_ChpxFkp.append( s_ChpxFkp )
            
            
#            print "\t" * 4 + "aFC\t\t\tPnFkpChpx\tStartChpxFkp"
#            for n_Index in range( l_PnFkpChpx.__len__() ) :
#                print "\t" * 4 + "0x%08X\t0x%08X\t0x%08X" % (l_FC[n_Index], l_PnFkpChpx[n_Index], l_StartChpxFkp[n_Index])
                        
            
        except :
            print format_exc()
            return None, None
        
        return l_FC_Text, l_ChpxFkp


    #    l_PnFkpChpx                                   [IN]                         PnFkpChpx List
    #    l_StartChpxFkp                                [OUT]                        ChpxFkp Start Offset List
    def fnParsePnFkpChpx(self, l_PnFkpChpx):
        
        try :
            
            l_StartChpxFkp = []
            
            for n_PnFkpChpx in l_PnFkpChpx :
                t_PnFkpChpx_Name = namedtuple("PnFkpChpx", RULE_PNFKPCHPX_NAME)
                s_PnFkpChpx = t_PnFkpChpx_Name._make( CBuffer.fnBitUnpack(RULE_PNFKPCHPX_PATTERN, n_PnFkpChpx) )
                l_StartChpxFkp.append( s_PnFkpChpx.Pn * 0x200 ) 
            
        except :
            print format_exc()
            return None
        
        return l_StartChpxFkp


    #    s_WordDoc                                    [IN]                            WordDocument Buffer
    #    l_ChpxFkp                                    [IN]                            ChpxFkp Buffer
    #    dl_Rgfc                                      [IN]                            Rgfc List ( Text Offset List Per Dump )
    #    dl_Rgb                                       [IN]                            Rgb List ( Text Property Offset List Per Dump)
    #    dl_Rgfc_Text                                 [OUT]                           Rgfc Member Dump List
    #    dl_Chpx                                      [OUT]                           Rgc Member Dump List ( Chpx )
    def fnParseChpx(self, s_WordDoc, l_ChpxFkp, dl_Rgfc, dl_Rgb) :
        
        try :
            
            # Dump Text referred Text Offset List
            dl_Rgfc_Text = []
            l_Rgfc_Text = []
            for n_DumpIndex in range( l_ChpxFkp.__len__() ) :
                
#                print "\t" * 4 + "[Rgfc] DumpIndex : 0x%08X" % n_DumpIndex
                
                for n_Index in range( dl_Rgfc[n_DumpIndex].__len__() ) :
                    if n_Index + 1 == dl_Rgfc[n_DumpIndex].__len__() :
                        break
                    n_Offset = dl_Rgfc[n_DumpIndex][n_Index]
                    n_Size = dl_Rgfc[n_DumpIndex][n_Index + 1] - dl_Rgfc[n_DumpIndex][n_Index]
                    s_Rgfc = s_WordDoc[n_Offset:n_Offset + n_Size]
                    l_Rgfc_Text.append( s_Rgfc )
                    
#                    print "\t" * 4 + "Index : 0x%08X - Offset : 0x%08X, Size : 0x%08X" % (n_Index, n_Offset, n_Size)
                    
                dl_Rgfc_Text.append( l_Rgfc_Text )
            

            # Dump Chpx referred Text Property Offset List
            dl_Chpx = []
            l_Chpx = []
            for n_DumpIndex in range( l_ChpxFkp.__len__() ) :
                
#                print "\t" * 4 + "[Rgb] DumpIndex : 0x%08X" % n_DumpIndex
                
                for n_Index in range( dl_Rgb[n_DumpIndex].__len__() ) :
                    n_Offset = dl_Rgb[n_DumpIndex][n_Index] * 2
                    if n_Offset == 0 :
                        continue
                    
                    n_cb = CBuffer.fnReadByte(l_ChpxFkp[n_DumpIndex], n_Offset)
                    n_Offset += 1
                    if n_cb == 0 :
                        continue
                    s_Chpx = l_ChpxFkp[n_DumpIndex][n_Offset:n_Offset + n_cb]
                    l_Chpx.append( s_Chpx )
                    
#                    print "\t" * 4 + "Index : 0x%08X - Offset : 0x%08X, Size : 0x%08X" % (n_Index, n_Offset, s_Chpx.__len__())
                    
                dl_Chpx.append( l_Chpx )

            
        except :
            print format_exc()
            return None, None
        
        return dl_Rgfc_Text, dl_Chpx


    #    dl_Chpx                                [IN]                    Double-List Chpx per Dump
    #    dl_Prl_Operand                         [OUT]                   Double-List Prl Operand
    def fnParseGrpPrl(self, dl_Chpx):
        
        try :
            
            MappedBasic = CMappedBasic()
            
            for n_DumpIndex in range( dl_Chpx.__len__() ) :
                l_Chpx = dl_Chpx[n_DumpIndex]
                B_Ret = MappedBasic.fnParsePrl( l_Chpx )
                if B_Ret != True :
                    print "\t" * 4 + "[-] Failure - ParsePrl( Chpx Dump 0x%08X nd )" % (n_DumpIndex)
                    return False
                
        except :
            print format_exc()
            return None
        
        return True
        


class CMappedBasic():
    
    #    s_Table                                          [IN]                        Table Stream Buffer ( 0Table or 1Table )
    #    n_Offset                                         [IN]                        Member Offset in Table Stream
    #    n_Size                                           [IN]                        Member Size in Table Stream
    #    l_Plc                                            [OUT]                       Plc.aFC Offset List
    #    l_Data                                           [OUT]                       Plc.aData Offset List
    def fnParsePlc(self, s_Table, n_Offset, n_Size):
        
        try :
            s_Plc = s_Table[n_Offset:n_Offset+n_Size]
            n_Cnt = ((n_Size - 4) / 4) / 2
                
            # Get Plc.aCP
            l_aCP = []
            n_tmpCnt = 0
            n_Position = 0
            while n_tmpCnt < n_Cnt :
                n_aCP = CBuffer.fnReadDword(s_Plc, n_Position)
                l_aCP.append( n_aCP )
                n_tmpCnt += 1
                n_Position += 4
                
            n_aCP = CBuffer.fnReadDword(s_Plc, n_Position)
            l_aCP.append( n_aCP )
            n_Position += 4
                
            # Get Plc.aData
            l_aData = []
            n_tmpCnt = 0
            while n_tmpCnt < n_Cnt :
                n_aData = CBuffer.fnReadDword(s_Plc, n_Position)
                l_aData.append( n_aData )
                n_tmpCnt += 1
                n_Position += 4
    
        except :
            print format_exc()
            return None, None
    
        return l_aCP, l_aData
    

    #    s_WordDoc                                      [IN]                        WordDocument Buffer    
    #    l_ParseDump                                    [IN]                        Parse Target Dump in WordDocument
    #    l_Member1                                      [OUT]                       ParsrTargetDump's Member1
    #    l_Member2                                      [OUT]                       ParsrTargetDump's Member2
    def fnParseCPwithCRUN(self, l_ParseDump):
        
        try :
            # Get CRUN
            l_CRUN = []
            for s_Dump in l_ParseDump :
                n_CRUN = CBuffer.fnConvertBinaryString2Int(s_Dump[-1], 1)
                if n_CRUN <= 0x0 or n_CRUN > 0x65 :
                    print "\t" * 5 + "[-] Error - ConvertBinaryString2Int( 0x%04X ) in ParseCPwithCRUN()" % n_CRUN
                    return None, None
                l_CRUN.append( n_CRUN )
            
            
            # Parse Target Dump List's Members
            dl_Member1 = []
            l_tmpMem1 = []
            dl_Member2 = []
            l_tmpMem2 = []
            for n_Index in range( l_ParseDump.__len__() ) :
                s_Dump = l_ParseDump[n_Index]
                n_Position = 0
                
                n_tmpCnt = 0
                while n_tmpCnt < l_CRUN[n_Index] + 1 :
                    l_tmpMem1.append( CBuffer.fnReadDword(s_Dump, n_Position) )
                    n_tmpCnt += 1
                    n_Position += 4
                dl_Member1.append( l_tmpMem1 )
                
                n_tmpCnt = 0
                while n_tmpCnt < l_CRUN[n_Index] :
                    l_tmpMem2.append( CBuffer.fnReadByte(s_Dump, n_Position) )
                    n_tmpCnt += 1
                    n_Position += 1
                dl_Member2.append( l_tmpMem2 )                
            
        except :
            print format_exc()
            return None, None
        
        return dl_Member1, dl_Member2
        
    
    #    l_Chpx                                        [IN]                        Chpx Buffer List
    #    l_Prl_Operand                                 [IN]                        Prl's Operand List
    def fnParsePrl(self, l_Chpx):
        
        try :
            
            for n_Index in range( l_Chpx.__len__() ) :
                s_Chpx = l_Chpx[n_Index]
                n_Sprm = CBuffer.fnReadWord(s_Chpx, 0)
                t_Sprm = self.fnParseSprm(n_Sprm)
                
                if t_Sprm.sgc > 0 and t_Sprm.sgc < 6 :
                    eval( "self.fnParse" + SINGLE_PROPERTY_MODIFIER[ t_Sprm.sgc ] + "PropertyModifier" )( n_Sprm, t_Sprm )
                else :
                    print "\t" * 3 + "[-] Error - 0x%04X" % n_Sprm,
                    print t_Sprm
                    continue
        
        except :
            print format_exc()
            return None

        return True
  
        
    #    n_Sprm                                        [IN]                        Sprm Data
    #    t_Sprm                                        [OUT]                       Mapped Sprm tuple
    def fnParseSprm(self, n_Sprm):
        
        try :
            
            t_Sprm_Name = namedtuple("Sprm", RULE_SPRM_NAME)
            t_Sprm = t_Sprm_Name._make( CBuffer.fnBitUnpack(RULE_SPRM_PATTERN, n_Sprm) )
            return t_Sprm
            
        except : 
            print format_exc()
            return None

    
    #    n_Sprm                                        [IN]                        Sprm Data
    #    t_Sprm                                        [IN]                        Mapped Sprm tuple
    #                                                  [OUT]                       True / False / None
    def fnParseCharPropertyModifier(self, n_Sprm, t_Sprm ):
        
        try : 
            
            print "\t" * 3 + "[*] Character Property : 0x%04X - %s" % (n_Sprm, self.fnCheckCharPropertyModifier(n_Sprm) )
            
#            CharPropertyModifier = CCharPropertyModifier()
#            eval( "CharPropertyModifier.fnParse" + self.fnCheckCharPropertyModifier(n_Sprm) )
        
        except :
            print format_exc()
            return None
        
        return True
        
    
    #
    def fnParseParaPropertyModifier(self, n_Sprm, t_Sprm ):
        
        try : 
        
            print "\t" * 3 + "[*] Paragraph Property : 0x%04X - %s" % (n_Sprm, self.fnCheckParaPropertyModifier(n_Sprm) )
            
#            ParagraphPropertyModifier = CParagraphPropertyModifier()
#            eval( "ParagraphPropertyModifier.fnParse" + self.fnCheckParaPropertyModifier(n_Sprm) )
        
        except :
            print format_exc()
            return None

        return True
    

    #
    def fnParseTablePropertyModifier(self, n_Sprm, t_Sprm ):
        
        try : 
        
            print "\t" * 3 + "[*] Table Property : 0x%04X - %s" % (n_Sprm, self.fnCheckTablePropertyModifier(n_Sprm) )
        
#            TablePropertyModifier = CTablePropertyModifier()
#            eval( "TablePropertyModifier.fnParse" + self.fnCheckTablePropertyModifier(n_Sprm) )
        
        except :
            print format_exc()
            return None
        
        return True


    #
    def fnParseSectPropertyModifier(self, n_Sprm, t_Sprm ):
        
        try : 
        
            print "\t" * 3 + "[*] Section Property : 0x%04X - %s" % (n_Sprm, self.fnCheckSectPropertyModifier(n_Sprm) )
            
#            SectionPropertyModifier = CSectionPropertyModifier()
#            eval( "SectionPropertyModifier.fnParse" + self.fnCheckSectPropertyModifier(n_Sprm) )
        
        except :
            print format_exc()
            return None
        
        return True
        

    #
    def fnParsePictPropertyModifier(self, n_Sprm, t_Sprm ):
        
        try : 
        
            print "\t" * 3 + "[*] Picture Property : 0x%04X - %s" % (n_Sprm, self.fnCheckPictPropertyModifier(n_Sprm) )
        
#            PicturePropertyModifier = CPicturePropertyModifier()
#            eval( "PicturePropertyModifier.fnParse" + self.fnCheckPictPropertyModifier(n_Sprm) )
        
        except :
            print format_exc()
            return None
        
        return True
        

    #    n_Sprm                                        [IN]                        Sprm Data
    #    s_CPM_Name                                    [OUT]                       Entity By n_Sprm in CHAR_PROPERTY_MODIFIER_NAME
    def fnCheckCharPropertyModifier(self, n_Sprm):
        
        try :
            
            s_CPM_Name = ""
            if n_Sprm in CHAR_PROPERTY_MODIFIER_VALUE :
                n_Index = CHAR_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_CPM_Name = CHAR_PROPERTY_MODIFIER_NAME[ n_Index ]
                
            return s_CPM_Name
        
        except :
            print format_exc()
            return None


    #    n_Sprm                                        [IN]                        Sprm Data
    #    s_PPM_Name                                    [OUT]                       Entity By n_Sprm in PARA_PROPERTY_MODIFIER_NAME
    def fnCheckParaPropertyModifier(self, n_Sprm):
        
        try :
            
            s_PPM_Name = ""
            if n_Sprm in PARA_PROPERTY_MODIFIER_VALUE :
                n_Index = PARA_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_PPM_Name = PARA_PROPERTY_MODIFIER_NAME[ n_Index ]
            
            return s_PPM_Name
        
        except :
            print format_exc()
            return None

    
    #    n_Sprm                                        [IN]                        Sprm Data
    #    s_TPM_Name                                    [OUT]                       Entity By n_Sprm in TABLE_PROPERTY_MODIFIER_NAME
    def fnCheckTablePropertyModifier(self, n_Sprm):
        
        try :
            # Referred page 136 in [MS-DOC].pdf
            s_TPM_Name = ""
            if n_Sprm in TABLE_PROPERTY_MODIFIER_VALUE :
                n_Index = TABLE_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_TPM_Name = TABLE_PROPERTY_MODIFIER_NAME[ n_Index ]
            
            return s_TPM_Name
                         
        except :
            print format_exc()
            return None


    #    n_Sprm                                        [IN]                        Sprm Data
    #    s_SPM_Name                                    [OUT]                       Entity By n_Sprm in SECT_PROPERTY_MODIFIER_NAME
    def fnCheckSectPropertyModifier(self, n_Sprm):
        
        try :
            
            s_SPM_Name = ""
            if n_Sprm in SECT_PROPERTY_MODIFIER_VALUE :
                n_Index = SECT_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_SPM_Name = SECT_PROPERTY_MODIFIER_NAME[ n_Index]
            
            return s_SPM_Name
        
        except :
            print format_exc()
            return None


    #    n_Sprm                                        [IN]                        Sprm Data
    #    s_PPM_Name                                    [OUT]                       Entity By n_Sprm in PICT_PROPERTY_MODIFIER_NAME
    def fnCheckPictPropertyModifier(self, n_Sprm):
        
        try :
            
            s_PPM_Name = ""
            if n_Sprm in PICT_PROPERTY_MODIFIER_VALUE :
                n_Index = PICT_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_PPM_Name = PICT_PROPERTY_MODIFIER_NAME[ n_Index ]
            
            return s_PPM_Name
        
        except :
            print format_exc()
            return None
        
        
        
        
        
        
        
        
        
            
    