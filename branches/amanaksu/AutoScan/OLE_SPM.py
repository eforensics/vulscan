######################################
#                                    #
#    Single Property Modifier        #
#                                    #
######################################

from traceback import format_exc


from Definition import *
from Common import CBuffer, CFile


class CSPM():
    
    def fnCheckCharPropertyModifier(self, n_Sprm):
        #    n_Sprm                                        [IN]                        Sprm Data
        #    s_CPM_Name                                    [OUT]                       Entity By n_Sprm in CHAR_PROPERTY_MODIFIER_NAME        
        try :
            
            s_CPM_Name = ""
            if n_Sprm in CHAR_PROPERTY_MODIFIER_VALUE :
                n_Index = CHAR_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_CPM_Name = CHAR_PROPERTY_MODIFIER_NAME[ n_Index ]
                
            return s_CPM_Name
        
        except :
            print format_exc()
            return None
    def fnParseCharPropertyModifier(self, n_Sprm, d_Sprm, s_Operand ):     
        #    n_Sprm               [IN]               Sprm Data
        #    d_Sprm               [IN]               Mapped Sprm tuple
        #                         [OUT]              True / False / None    
        try : 
            
            s_FuncName = self.fnCheckCharPropertyModifier(n_Sprm)
#            print "\t" * 4 + "[*] Character Property : 0x%04X %s - %s" % (n_Sprm, str(d_Sprm), s_FuncName )
            
            if s_FuncName != "" :
                CharPropertyModifier = CCharPropertyModifier()
#                if not eval( "CharPropertyModifier.fnParse" + s_FuncName )(d_Sprm, s_Operand) :
#                    print "\t" * 4 + "[-] Error ( %s )" % s_FuncName
#                    return False
        
        except :
            print format_exc()
            return None
        
        return True
    def fnCheckParaPropertyModifier(self, n_Sprm):
        #    n_Sprm               [IN]               Sprm Data
        #    s_PPM_Name           [OUT]              Entity By n_Sprm in PARA_PROPERTY_MODIFIER_NAME        
        try :
            
            s_PPM_Name = ""
            if n_Sprm in PARA_PROPERTY_MODIFIER_VALUE :
                n_Index = PARA_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_PPM_Name = PARA_PROPERTY_MODIFIER_NAME[ n_Index ]
            
            return s_PPM_Name
        
        except :
            print format_exc()
            return None 
    def fnParseParaPropertyModifier(self, n_Sprm, d_Sprm, s_Operand ):
        #    n_Sprm               [IN]               Sprm Data
        #    d_Sprm               [IN]               Mapped Sprm tuple
        #                         [OUT]              True / False / None        
        try : 
            
            s_FuncName = self.fnCheckParaPropertyModifier(n_Sprm)
#            print "\t" * 4 + "[*] Paragraph Property : 0x%04X %s - %s" % (n_Sprm, str(d_Sprm), s_FuncName )
            
            if s_FuncName != "" :
                ParagraphPropertyModifier = CParagraphPropertyModifier()
                if not eval( "ParagraphPropertyModifier.fnParse" + s_FuncName )(d_Sprm, s_Operand) :
                    print "\t" * 4 + "[-] Failure - %s" % s_FuncName
                    return False
        
        except :
            print format_exc()
            return None

        return True
    def fnCheckTablePropertyModifier(self, n_Sprm):
        #    n_Sprm               [IN]               Sprm Data
        #    s_PPM_Name           [OUT]              Entity By n_Sprm in PARA_PROPERTY_MODIFIER_NAME    
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
    def fnParseTablePropertyModifier(self, n_Sprm, d_Sprm, s_Operand ):
        #    n_Sprm               [IN]               Sprm Data
        #    d_Sprm               [IN]               Mapped Sprm tuple
        #                         [OUT]              True / False / None    
  
        try : 
        
            s_FuncName = self.fnCheckTablePropertyModifier(n_Sprm)
#            print "\t" * 4 + "[*] Table Property : 0x%04X %s - %s" % (n_Sprm, str(d_Sprm), s_FuncName )
        
            if s_FuncName != "" :
                TablePropertyModifier = CTablePropertyModifier()
                if not eval( "TablePropertyModifier.fnParse" + s_FuncName )(d_Sprm, s_Operand) :
                    print "\t" * 4 + "[-] Failure - %s" % s_FuncName
                    return False

        except :
            print format_exc()
            return None
        
        return True
    def fnCheckSectPropertyModifier(self, n_Sprm):
        #    n_Sprm               [IN]               Sprm Data
        #    s_PPM_Name           [OUT]              Entity By n_Sprm in PARA_PROPERTY_MODIFIER_NAME        
        try :
            
            s_SPM_Name = ""
            if n_Sprm in SECT_PROPERTY_MODIFIER_VALUE :
                n_Index = SECT_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_SPM_Name = SECT_PROPERTY_MODIFIER_NAME[ n_Index]
            
            return s_SPM_Name
        
        except :
            print format_exc()            
            return None
    def fnParseSectPropertyModifier(self, n_Sprm, d_Sprm, s_Operand ):
        #    n_Sprm               [IN]               Sprm Data
        #    d_Sprm               [IN]               Mapped Sprm tuple
        #                         [OUT]              True / False / None          
        try : 
        
            s_FuncName = self.fnCheckSectPropertyModifier(n_Sprm)
#            print "\t" * 4 + "[*] Section Property : 0x%04X %s - %s" % (n_Sprm, str(d_Sprm), s_FuncName )
            
            if s_FuncName != "" :
                SectionPropertyModifier = CSectionPropertyModifier()
                if not eval( "SectionPropertyModifier.fnParse" + s_FuncName )(d_Sprm, s_Operand) :
                    print "\t" * 4 + "[-] Failure - %s" % s_FuncName
                    return False
            
            
        except :
            print format_exc()
            return None
        
        return True
    def fnCheckPictPropertyModifier(self, n_Sprm):
        #    n_Sprm               [IN]               Sprm Data
        #    s_PPM_Name           [OUT]              Entity By n_Sprm in PARA_PROPERTY_MODIFIER_NAME        
        try :
            
            s_PPM_Name = ""
            if n_Sprm in PICT_PROPERTY_MODIFIER_VALUE :
                n_Index = PICT_PROPERTY_MODIFIER_VALUE.index( n_Sprm )
                s_PPM_Name = PICT_PROPERTY_MODIFIER_NAME[ n_Index ]
            
            return s_PPM_Name
        
        except :
            print format_exc()
            return None
    def fnParsePictPropertyModifier(self, n_Sprm, d_Sprm, s_Operand ):
        
        try : 
            
            s_FuncName = self.fnCheckPictPropertyModifier(n_Sprm)
#            print "\t" * 4 + "[*] Picture Property : 0x%04X %s - %s" % (n_Sprm, str(d_Sprm), s_FuncName )        
        
            if s_FuncName != "" :
                PicturePropertyModifier = CPicturePropertyModifier()
                if not eval( "PicturePropertyModifier.fnParse" + s_FuncName )(d_Sprm, s_Operand) :
                    print "\t" * 4 + "[-] Failure - %s" % s_FuncName
                    return False


        except :
            print format_exc()
            return None
        
        return True



class CCharPropertyModifier():
    
    def fnParsesprmCFRMarkDel(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFRMarkIns(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFFldVanish(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmCPicLocation(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmCIbstRMark(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCDttmRMark(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFData(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCIdslRMark(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCSymbol(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCFOle2(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCHighlight(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFWebHidden(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmCRsidProp(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCRsidText(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCRsidRMDel(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFSpecVanish(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmCFMathPr(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCIstd(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmCIstdPermute(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmCPlain(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCKcd(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCFBold(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCFItalic(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCFStrike(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCFOutline(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCFShadow(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCFSmallCaps(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmCFCaps(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCFVanish(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCKul(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCDxzSpace(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCIco(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCHps(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCHpsPos(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCMajority(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCIss(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCHpsKern(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCHresi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCRgFtc0(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCRgFtc1(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCRgFtc2(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCCharScale(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFDStrike(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCFImprint(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCFSpec(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCFObj(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmCPropRMark90(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmCFEmboss(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCSfxText(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCFBiDi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCFBoldBi(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCFItalicBi(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFtcBi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmClidBi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCIcoBi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCHpsBi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCDispFldRMark(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmCIbstRMarkDel(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmCDttmRMarkDel(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmCBrc80(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCShd80(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmCIdslRMarkDel(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmCFUsePgsuSettings(self, d_Sprm, s_Operand):              pass
    def fnParsesprmCRgLid0_80(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmRgLid1_80(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCIdctHint(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCCv(self, d_Sprm, s_Operand):                            pass
    def fnParsesprmCShd(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCBrc(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCRgLid0(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCRgLid1(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCFNoProof(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCFitText(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmCCvUl(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmCFELayout(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmCLbcCRJ(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCFComplexScripts(self, d_Sprm, s_Operand):               pass
    def fnParsesprmCWall(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmCCnf(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmCNeedFontFixup(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmCPbiIBullet(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmCPbiGrf(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmCPropRMark(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmCFSdtVanish(self, d_Sprm, s_Operand):                    pass


class CParagraphPropertyModifier():
    
    def fnParsesprmPIstd(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPIstdPermute(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPIncLvl(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmPJc80(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPFKeep(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmPFKeepFollow(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPFPageBreakBefore(self, d_Sprm, s_Operand):              pass
    def fnParsesprmPIlvl(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPIlfo(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPFNoLineNumb(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPChgTabsPapx(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPDxaRight80(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPDxaLeft80(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPNest80(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmPDxaLeft180(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPDyaLine(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmPDyaBefore(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPDyaAfter(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPChgTabs(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmPFInTable(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPFTtp(self, d_Sprm, s_Operand):                          pass                
    def fnParsesprmPDxaAbs(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmPDyaAbs(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmPDxaWidth(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPPc(self, d_Sprm, s_Operand):                            pass
    def fnParsesprmPWr(self, d_Sprm, s_Operand):                            pass
    def fnParsesprmPBrcTop80(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPBrcLeft80(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPBrcBottom80(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPBrcRight80(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPBrcBetween80(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmPBrcBar80(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPFNoAutoHyph(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPWHeightAbs(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPDcs(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmPShd80(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmPDyaFromText(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPDxaFromText(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPFLocked(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmPFWidowControl(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmPFKinsoku(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPFWordWrap(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPDOverflowPunct(self, d_Sprm, s_Operand):                pass
    def fnParsesprmPFTopLinePunct(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmPFAutoSpaceDE(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmPFAutoSpaceDN(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmPWAlignFont(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPFrameTextFlow(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmPOutLvl(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmPFBiDi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmPFNumRMIns(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPNumRM(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmPHugePapx(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPFUsePgsuSettings(self, d_Sprm, s_Operand):              pass
    def fnParsesprmPFAdjustRight(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmPItap(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPDtap(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPFInnerTableCell(self, d_Sprm, s_Operand):               pass
    def fnParsesprmPFInnerTtp(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPShd(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmPBrcTop(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmPBrcLeft(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmPBrcBottom(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPBrcRight(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPBrcBetween(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPBrcBar(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmPDxcRight(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPDxcLeft(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmPDxcLeft1(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPDylBefore(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPDylAfter(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPFOpenTch(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPFDyaBeforeAuto(self, d_Sprm, s_Operand):                pass
    def fnParsezsprmPFDyaAfterAuto(self, d_Sprm, s_Operand):                pass
    def fnParsesprmPDxaRight(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPDxaLeft(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmPNest(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPDxaLeft1(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPJc(self, d_Sprm, s_Operand):                            pass
    def fnParsesprmPFNoAllowOverlap(self, d_Sprm, s_Operand):               pass
    def fnParsesprmPWall(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPIpgp(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPCnf(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmPRsid(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmPIstdListPermute(self, d_Sprm, s_Operand):               pass
    def fnParsesprmPTableProps(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPTIstdInfo(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPFContextualSpacing(self, d_Sprm, s_Operand):            pass
    def fnParsesprmPPropRMark(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPFMirrorIndents(self, d_Sprm, s_Operand):                pass
    def fnParsesprmPTtwo(self, d_Sprm, s_Operand):                          pass
    

class CTablePropertyModifier():
    
    def fnParsesprmTJc90(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmTDxaLeft(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmTDxaGapHalf(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmTFCantSplit90(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmTTableHeader(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTTableBorders80(self, d_Sprm, s_Operand):                pass
    def fnParsesprmTDyaRowHeight(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmTDefTable(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmTDefTableShd80(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmTTlp(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmTFBiDi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmTDefTableShd3rd(self, d_Sprm, s_Operand):                pass
    def fnParsesprmTPc(self, d_Sprm, s_Operand):                            pass
    def fnParsesprmTDxaAbs(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmTDyaAbs(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmTDxaFromText(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTDyaFromText(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTDefTableShd(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTTableBorders(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmTTableWidth(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmTFAutofit(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmTDefTableShd2nd(self, d_Sprm, s_Operand):                pass
    def fnParsesprmTWidthBefore(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTWidthAfter(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmTFKeepFollow(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTBrcTopCv(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmTBrcLeftCv(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmTBrcBottomCv(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTBrcRightCv(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmTDxaFromTextRight(self, d_Sprm, s_Operand):              pass
    def fnParsesprmTDyaFromTextBottom(self, d_Sprm, s_Operand):             pass
    def fnParsesprmTSetBrc80(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmTInsert(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmTDelete(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmTDxaCol(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmTMerge(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmTSplit(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmTTextFlow(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmTVertMerge(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmTVertAlign(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmTSetShd(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmTSetShdOdd(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmTSetBrc(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmTCellPadding(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTCellSpacingDefault(self, d_Sprm, s_Operand):            pass
    def fnParsesprmTCellPaddingDefault(self, d_Sprm, s_Operand):            pass
    def fnParsesprmTCellWidth(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmTFitText(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmTFCellNoWrap(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTIstd(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmTCellPaddingStyle(self, d_Sprm, s_Operand):              pass
    def fnParsesprmTCellFHideMark(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmTSetShdTable(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTWidthIndent(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTCellBrcType(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmTFBiDi90(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmTFNoAllowOverlap(self, d_Sprm, s_Operand):               pass
    def fnParsesprmTFCantSplit(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmTPropRMark(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmTWall(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmTIpgp(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmTCnf(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmTDefTableShdRaw(self, d_Sprm, s_Operand):                pass
    def fnParsesprmTDefTableShdRaw2nd(self, d_Sprm, s_Operand):             pass
    def fnParsesprmTDefTableShdRaw3rd(self, d_Sprm, s_Operand):             pass
    def fnParsesprmTRsid(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmTCellVertAlignStyle(self, d_Sprm, s_Operand):            pass
    def fnParsesprmTCellNoWrapStyle(self, d_Sprm, s_Operand):               pass
    def fnParsesprmTCellBrcTopStyle(self, d_Sprm, s_Operand):               pass
    def fnParsesprmTCellBrcBottomStyle(self, d_Sprm, s_Operand):            pass
    def fnParsesprmTCellBrcLeftStyle(self, d_Sprm, s_Operand):              pass
    def fnParsesprmTCellBrcRightStyle(self, d_Sprm, s_Operand):             pass
    def fnParsesprmTCellBrcInsideHStyle(self, d_Sprm, s_Operand):           pass
    def fnParsesprmTCellBrcInsideVStyle(self, d_Sprm, s_Operand):           pass
    def fnParsesprmTCellBrcTL2BRStyle(self, d_Sprm, s_Operand):             pass
    def fnParsesprmTCellBrcTR2BLStyle(self, d_Sprm, s_Operand):             pass
    def fnParsesprmTCellShdStyle(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmTCHorzBands(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmTCVertBands(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmTJc(self, d_Sprm, s_Operand):                            pass


class CSectionPropertyModifier():
    
    def fnParsesprmScnsPgn(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSiHeadingPgn(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmSDxaColWidth(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmSDxaColSpacing(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmSFEvenlySpaced(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmSFProtected(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSDmBinFirst(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSDmBinOther(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSDkc(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmSFTitlePage(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSCcolumns(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmSDxaColumns(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSNfcPgn(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSFPgnRestart(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmSFEndnote(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmSLnc(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmSNLnnMod(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmSDxaLnn(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSDyaHdrTop(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSDyaHdrBottom(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmSLBetween(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmSVjc(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmSLnnMin(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSPgnStart97(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSBOrientation(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmSXaPage(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSYaPage(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSDxaLeft(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmSDxaRight(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmSDyaTop(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSDyaBottom(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSDzaGutter(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSDmPaperReq(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSFBiDi(self, d_Sprm, s_Operand):                         pass
    def fnParsesprmSFRTLGutter(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSBrcTop80(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmSBrcLeft80(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSBrcBottom80(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmSBrcRight80(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmSPgbProp(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmSDxtCharSpace(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmSDyaLinePitch(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmSClm(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmSTextFlow(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmSBrcTop(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSBrcLeft(self, d_Sprm, s_Operand):                       pass
    def fnParsesprmSBrcBottom(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSBrcRight(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmSWall(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmSRsid(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmSFpc(self, d_Sprm, s_Operand):                           pass
    def fnParsesprmSRncFtn(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSRncEdn(self, d_Sprm, s_Operand):                        pass
    def fnParsesprmSNFtn(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmSNfcFtnRef(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSNEdn(self, d_Sprm, s_Operand):                          pass
    def fnParsesprmSNfcEdnRef(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSPropRMark(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmSPgnStart(self, d_Sprm, s_Operand):                      pass


class CPicturePropertyModifier():
    
    def fnParsesprmPicBrcTop80(self, d_Sprm, s_Operand):                    pass
    def fnParsesprmPicBrcLeft80(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPicBrcBottom80(self, d_Sprm, s_Operand):                 pass
    def fnParsesprmPicBrcRight80(self, d_Sprm, s_Operand):                  pass
    def fnParsesprmPicBrcTop(self, d_Sprm, s_Operand):                      pass
    def fnParsesprmPicBrcLeft(self, d_Sprm, s_Operand):                     pass
    def fnParsesprmPicBrcBottom(self, d_Sprm, s_Operand):                   pass
    def fnParsesprmPicBrcRight(self, d_Sprm, s_Operand):                    pass



