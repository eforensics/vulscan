
import os, binascii
from struct import unpack, error
from traceback import format_exc
from collections import namedtuple


from Definition import *
from Common import CBuffer, CFile, CConvertString
from OLE_SPM import CSPM
from OLE_Cmds import CCMDS
from PE import CPE




class CStreamDOC():

######################
#                    #
#    Office97        #
#                    #
######################

    #    s_fname               [IN]               File Full Name
    #    s_WordDoc             [IN]               WordDocument Stream Buffer
    #    s_Table               [IN]               Table Stream Buffer ( 0Table or 1Table )
    #    n_Offset              [IN]               Member Offset in Table Stream
    #    n_Size                [IN]               Member Size in Table Stream
    #    BOOL Type             [OUT]              True / False    
    
    def fnStshf(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        #    Stshf    : Style Sheet
        #    Referred Table Stream 
        print "\t" * 3 + "[+] Stshf()"
        
        try :
            
            if n_Size == 0 :
                print "\t" * 4 + "[-] Error - \"Stshf\" Size MUST be a nonzero value"
                return False
            
            CFile.fnWriteFile("%s\\%s.Stshf" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnPlcfSed(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfSed()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.PlcfSed" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
#            MappedBasic = CMappedBasic()
#            
## Step 1. Mapped Plc
#            l_fc, l_Data = MappedBasic.fnMappedPlc(s_Table, n_Offset, n_Size, 0xc, 12)
#            if l_fc == None or l_Data == None :
#                print "\t" * 3 + "[-] Failure - MappedPlc()"
#                return False
            
        except :
            print format_exc()
            return False
        
        return True

    def fnPlcfHdd(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfHdd()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.PlcfHdd" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
                
    def fnPlcBteChpx(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcBteChpx()"
        
        try :
            CFile.fnWriteFile("%s\\%s.PlcBteChpx" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
            MappedBasic = CMappedBasic()
            MappedStream = CMappedStream()
            
# Step 1. Mapped Plc - OK
            l_fc, l_Pn = MappedBasic.fnMappedPlc(s_Table, n_Offset, n_Size, None, 4)
            if l_fc == None or l_Pn == None :
                print "\t" * 3 + "[-] Failure - MappedPlc()"
                return False
            
# Extra Step. Check General Vulnerability
            Vulnerability = CVulnerability()
            for n_Index in range( l_fc.__len__() - 1 ) :
                if not Vulnerability.fnGeneralPE(s_WordDoc[l_fc[n_Index]:l_fc[n_Index+1]-l_fc[n_Index]]) :
                    print "\t" * 4 + "[-] Failure - General()"
                    
            
# Step 2. Mapped Property List - OK
            l_Fkp, dl_Rgfc, dl_Rgb = MappedBasic.fnMappedCP(s_WordDoc, l_Pn, 0x65, 0x1, 1)
            if l_Fkp == None or dl_Rgfc == None or dl_Rgb == None :
                print "\t" * 3 + "[-] Failure - MappedCP()"
                return False
            if l_Fkp == [] or dl_Rgfc == [] or dl_Rgb == [] :
                return True
            
            
# Step 3. Mapped GroupPrl 
            dl_Group = MappedStream.fnMappedGrpPrl(l_Fkp, dl_Rgb)
            if dl_Group == None :
                print "\t" * 3 + "[-] Failure - MappedGrpPrl()"
                return False


# Step 4. Mapped Prl
            for l_GrpPrl in dl_Group :
                if not MappedStream.fnMappedPrl(l_GrpPrl) :
                    print "\t" * 3 + "[-] Failure - MappedPrl()"
                    return False

            PrintDoc = CPrintDoc()
            s_OutPrint = "< PlcBteChpx - Referred TableStream >\n"
            s_OutPrint += PrintDoc.fnPlcBteChpx(s_WordDoc, s_Table, l_fc, l_Pn, l_Fkp, dl_Rgfc, dl_Rgb, dl_Group)
            with open( "%s\\%s_PlcBteChpx.log" % (g_DumpPath, s_fname), "wb" ) as p_file :
                p_file.write( s_OutPrint )

        except :
            print format_exc()
            return False
        
        return True
     
    def fnPlcfBtePapx(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfBtePapx()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.PlcBtePapx" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
            MappedBasic = CMappedBasic()
            MappedStream = CMappedStream()
            
# Step 1. Mapped Plc
            l_fc, l_Pn = MappedBasic.fnMappedPlc(s_Table, n_Offset, n_Size, None, 4)
            if l_fc == None or l_Pn == None :
                print "\t" * 3 + "[-] Failure - MappedPlc()"
                return False
             
            
# Step 2. Mapped Property List
            l_Fkp, dl_Rgfc, dl_Rgbx = MappedBasic.fnMappedCP(s_WordDoc, l_Pn, 0x1D, 0x1, 0xD)
            if l_Fkp == None or dl_Rgfc == None or dl_Rgbx == None :
                print "\t" * 3 + "[-] Failure - MappedCP()"
                return False  
            if l_Fkp == [] or dl_Rgfc == [] or dl_Rgbx == [] :
                return True
            
            
# Step 3. Mapped GroupPrlIstd
            dl_istd, dl_Group = MappedStream.fnMappedGrpPrlIstd(l_Fkp, dl_Rgbx)
            if dl_istd == None or dl_Group == None :
                print "\t" * 3 + "[-] Failure - MappedGrpPrl()"
                return False
            
            
# Step 4. Mapped Prl
            for l_GrpPrl in dl_Group :
                if not MappedStream.fnMappedPrl(l_GrpPrl) :
                    print "\t" * 3 + "[-] Failure - MappedPrl()"
                    return False
            
#            PrintDoc = CPrintDoc()
#            s_OutPrint = "< PlcfBtePapx - Referred TableStream >\n"
#            s_OutPrint += PrintDoc.fnPlcfBtePapx(s_WordDoc, s_Table, l_fc, l_Pn, l_Fkp, dl_Rgfc, dl_Rgbx, dl_istd, dl_Group)
#            with open( "%s\\%s_PlcfBtePapx.log" % (g_DumpPath, s_fname), "wb" ) as p_file :
#                p_file.write( s_OutPrint )
            
        except :
            print format_exc()
            return False
        
        return True        

    def fnSttbfFfn(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] SttbfFfn()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.SttbfFfn" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True     

    def fnCmds(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] Cmds()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.Cmds" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
            MappedStream = CMappedStream()
            
            s_Tcg = s_Table[n_Offset:n_Offset + n_Size]
            
            n_Position = 0
            nTcgVer = CBuffer.fnReadByte(s_Tcg, n_Position)
            n_Position += 1
            
            if nTcgVer != 0xFF :        # MUST be 255 ( 0xFF )
                print "\t" * 4 + "[-] Error ( Tcg Version : 0x%02X )" % nTcgVer
                return False
            
            if not MappedStream.fnMappedTcg255(s_Tcg[1:]) :
                print "\t" * 4 + "[-] Failure - ParseTcg255()"
                return False
            
        except :
            print format_exc()
            return False
        
        return True

    def fnWss(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] Wss()"
        
        try :
            
            return
            
        except :
            print format_exc()
            return False
        
        return True

    def fnDop(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] Dop()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.Dop" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnSttbfAssoc(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] SttbfAssoc()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.SttbfAssoc" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnClx(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] Clx()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.Clx" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
      
    def fnDggInfo(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] DggInfo()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.DggInfo" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True

    def fnSttbfRMark(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] SttbfRMark()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.SttbfRMark" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnPlcfSpl(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfSpl()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.PlcfSpl" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnPlcfGram(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfGram()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.PlcfGram" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
    
######################
#                    #
#    Office2000      #
#                    #
######################

    def fnPlcfTch(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfTch()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.PlcfTch" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnRmdThreading(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] RmdThreading()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.RmdThreading" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True    

    def fnPlcfLad(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfLad()"
        
        try :
            
            return True
            
        except :
            print format_exc()
            return False
        
        return True 

######################
#                    #
#    Office2002      #
#                    #
######################

    def fnSttbfBkmkFactoid(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] SttbfBkmkFactoid()"
        
        try :
            
            MappedStream = CMappedStream()
            MappedBasic = CMappedBasic()
            
            s_SttbfBkmkFactoid = s_Table[n_Offset:n_Offset + n_Size]
            
            CFile.fnWriteFile("%s\\%s.SttbfBkmkFactoid" % (g_DumpPath, s_fname), s_SttbfBkmkFactoid)
            
# Step 1. Mapped fExtend, cData, cbExtra, cchData0
            d_SttbHdr = MappedStream.fnMappedSttbHdr(s_SttbfBkmkFactoid)
            if d_SttbHdr == None :
                print "\t" * 4 + "[-] Failure - MappedSttbHdr()"
                return False
            

# Step 2. Mapped FactoidInfo
            l_FactoidInfo = []
            n_Cnt = d_SttbHdr["cData"]
            n_Position = SIZE_OF_STTBHDR
            while True :
                l_Info = MappedBasic.fnMappedFACTOIDINFO(s_SttbfBkmkFactoid[n_Position:n_Position+SIZE_OF_FACTOIDINFO])
                l_FactoidInfo.append( l_Info )
                
                # Enable Vulnerability
                Vulnerability = CVulnerability()
                if not Vulnerability.fnCVE200062492(l_Info) :
                    print "\t" * 4 + "[-] Failure - CVE20062492( %s )" % l_Info
                
                n_Position += SIZE_OF_FACTOIDINFO
                n_Cnt -= 1
                if n_Cnt == 0 or n_Position == (s_SttbfBkmkFactoid.__len__() - SIZE_OF_STTBHDR) :
                    break 
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnPlcfBklFactoid(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfBklFactoid()"

        try :
            
            MappedStream = CMappedStream()
            
# Step 1. Mapped FBKLD
            l_cp, l_Fbkld = MappedStream.fnMappedFbkld(s_pBuf, l_StartOffset, s_Table, n_Offset, n_Size)
            if l_cp == None or l_Fbkld == None :
                print "\t" * 4 + "[-] Failure - MappedFbkld()"
                return False
            
        except :
            print format_exc()
            return False
        
        return True

    def fnPlcfBkfFactoid(self, s_fname, s_pBuf, s_WordDoc, l_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfBkfFactoid()"
        
        try : 
            
            MappedStream = CMappedStream()
            
# Step 1. Mapped FBKFD
            l_cp, l_Fbkfd = MappedStream.fnMappedFbkfd(s_pBuf, l_StartOffset, s_Table, n_Offset, n_Size)
            if l_cp == None or l_Fbkfd == None :
                print "\t" * 4 + "[-] Failure - MappedFbkfd()"
                return False
            
            
        except :
            print format_exc()
            return False
        
        return True

    def fnFactoidData(self, s_fname, s_pBuf, s_WordDoc, d_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] FactoidData()"
    
        try :
            
            MappedStream = CMappedStream()
            
            s_SmartTagData = s_Table[n_Offset:n_Offset + n_Size]
            
# Step 1. Mapped PropBagStore
            dl_PropBagStore, n_Position = MappedStream.fnMappedPropBagStore(s_SmartTagData)
            
            
# Step 2. Mapped PropBags            
            dl_PropBags = MappedStream.fnMappedPropBags(s_SmartTagData[n_Position:])
            
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnPlcfPgp(self, s_fname, s_pBuf, s_WordDoc, d_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] PlcfPgp()"
        
        try :
            
            return True
            
        except :
            print format_exc()
            return False
        
        return True 

    def fnPlrsid(self, s_fname, s_pBuf, s_WordDoc, d_StartOffset, s_Table, n_Offset, n_Size):
        
        print "\t" * 3 + "[+] Plrsid()"
        
        try :
            
            CFile.fnWriteFile("%s\\%s.Plrsid" % (g_DumpPath, s_fname), s_Table[n_Offset:n_Offset + n_Size])
            
        except :
            print format_exc()
            return False
        
        return True 


######################
#                    #
#    Office2003      #
#                    #
######################


class CMappedStream():

    def fnMappedTcg255(self, s_tcg):
        
        try :
            
            CMDS = CCMDS()
            
            n_Position = 0
            n_Id = CBuffer.fnReadByte(s_tcg, n_Position)
            n_Position += 1
            
            if n_Id in CMDS_ID_FUNC.keys() :
                eval( "CMDS.fn%s" % CMDS_ID_FUNC[n_Id] )()
            elif n_Id == 0x40 :
                print "\t" * 4 + "[*] Terminate Sequence!"
            else :
                print "\t" * 4 + "[-] Error ( Tcg255's Id : 0x%02X [Undefined] )" % n_Id; return False
            
        except :
            print format_exc()
            return None
        
        return True

    def fnMappedGrpPrl(self, l_Fkp, dl_Rg):
        
        try :
            
            dl_Group = []
            l_Group = []
            
            for n_Index in range( l_Fkp.__len__() ) :
                s_Fkp = l_Fkp[n_Index]
                l_Rg = dl_Rg[n_Index]
                
                for n_Offset in l_Rg :
                    n_Offset = n_Offset * 2
                    n_cb = CBuffer.fnReadByte(s_Fkp, n_Offset)
                    l_Group.append( s_Fkp[n_Offset + 1:n_Offset + 1 + n_cb])
            
                dl_Group.append( l_Group )
            
        except :
            print format_exc()
            return None
        
        return dl_Group

    def fnMappedGrpPrlIstd(self, l_Fkp, dl_Rg):
        
        try :
            
            Check = CCheck()
            
            dl_Group = []
            dl_istd = []
            l_Group = []
            l_istd = []

            for n_Index in range( l_Fkp.__len__() ) :
                s_Fkp = l_Fkp[n_Index]
                l_Rg = dl_Rg[n_Index]
                
                l_NonDupRg = Check.fnChckDuplicateList( l_Rg )
                
                for s_Rg in l_NonDupRg :
                    n_Offset = CBuffer.fnReadByte(s_Rg, 0) * 2
                    n_cb = CBuffer.fnReadByte(s_Fkp, n_Offset)
                    n_Offset += 1
                    if n_cb == 0 :
                        n_cb = CBuffer.fnReadByte(s_Fkp, n_Offset)
                        n_Offset += 1
                        n_Size = n_cb * 2
                        
                    else :
                        n_Size = n_cb * 2 -1
                    
                    l_istd.append( CBuffer.fnReadWord(s_Fkp, n_Offset) )
                    n_Offset += 2
                    l_Group.append( s_Fkp[n_Offset:n_Offset + n_Size-2] )
                    
                dl_Group.append( l_Group )
                dl_istd.append( l_istd )
                
        except :
            print format_exc()
            return None

        return dl_istd, dl_Group

    def fnMappedPrl(self, l_GrpPrl):
        #    l_GrpPrl                        [IN]                    Grouped Prl List
        #                                    [OUT]                   True / None
        try :
            
            MappedBasic = CMappedBasic()
            SPM = CSPM()
            
            for s_GrpPrl in l_GrpPrl :
                if s_GrpPrl == "" :
                        continue
                
                n_Position = 0
                while True :
                    n_Sprm = CBuffer.fnReadWord(s_GrpPrl, n_Position)
                    n_Position += 2
                    
                    n_ispmd = n_Sprm & 0x01FF
                    n_f = (n_Sprm / 0x200) & 0x0001
                    n_sgc = (n_Sprm / 0x400) & 0x0007
                    n_spra = n_Sprm / 0x2000
                    
                    d_Sprm = {"ispmd":n_ispmd, "fSpec":n_f, "sgc":n_sgc, "spra":n_spra}
                    
                    if n_spra == 0 :           n_szOperand = 1     
                    elif n_spra == 1 :         n_szOperand = 1     
                    elif n_spra == 2 :         n_szOperand = 2     
                    elif n_spra == 3 :         n_szOperand = 4     
                    elif n_spra == 4 :         n_szOperand = 2     
                    elif n_spra == 5 :         n_szOperand = 2     
                    elif n_spra == 6 :         n_szOperand = CBuffer.fnReadByte(s_GrpPrl, n_Position);  n_Position += 1
                    elif n_spra == 7 :         n_szOperand = 3     
                    else :
                        print "\t" * 4 + "[-] Error ( Spra : 0x%04X )" % n_spra
                        return None, None, None
                    
                    s_Operand = s_GrpPrl[n_Position:n_Position + n_szOperand]
                    n_Position += n_szOperand
                    if n_Position == s_GrpPrl.__len__() or n_Position > s_GrpPrl.__len__() :
                        break
                    
#                    if n_sgc > 0 and n_sgc < 6 :
#                        eval( "SPM.fnParse" + SINGLE_PROPERTY_MODIFIER[ n_sgc ] + "PropertyModifier" )( n_Sprm, d_Sprm, s_Operand )
#                        if not eval( "SPM.fnParse" + SINGLE_PROPERTY_MODIFIER[ n_sgc ] + "PropertyModifier" )( n_Sprm, d_Sprm, s_Operand ) :
#                            print "\t" * 4 + "%sPropertyModifier()" % SINGLE_PROPERTY_MODIFIER[ n_sgc ]    
#                    else :
#                        print "\t" * 4 + "[-] %s" % str(d_Sprm)
                
        except :
            print format_exc()
            return None
        
        return True

    def fnMappedPropBagStore(self, s_SmartTagData):
        
        try :
            
            MappedBasic = CMappedBasic()
            
            dl_PropBagStore = []
            
# Step 1. Get "propBagStore.cFactoidType" - Array factoidTypes's Element Count 
            n_Position = 0
            n_cFactoidType = CBuffer.fnReadDword(s_SmartTagData, n_Position)
            dl_PropBagStore.append( n_cFactoidType )
            n_Position += 4
            
            
# Step 2. Get "propBagStore.factoidTypes" - Array factoidTypes
            n_Cnt = 0
            while True :
                l_FactoidType, n_Position = self.fnMappedFactoidType(s_SmartTagData, n_Position)
                dl_PropBagStore.append( l_FactoidType )
                n_Cnt += 1
                if n_Cnt == n_cFactoidType :
                    break
            

# Step 3. Get "propBagStore.hdr ( cbHdr, sVer, cFactoid, cste )"
            t_Hdr = MappedBasic.fnMappedPropBagStoreHdr(s_SmartTagData[n_Position:n_Position + 0xC])
            if t_Hdr == None :
                print "\t" * 4 + "[-] Failure - MappedPropBagStoreHdr()"
                return None, None
            dl_PropBagStore.append( t_Hdr )
            n_Position += t_Hdr.cbHdr
            
            
# Step 4. Mapped stringTable
            l_stringTable, n_Offset = MappedBasic.fnMappedPBStrings(s_SmartTagData[n_Position:], t_Hdr.cste)
            if l_stringTable == None or n_Offset == None :
                print "\t" * 4 + "[-] Failure - MappedPBStrings( Buffer : StringTable, Position : 0x%08X, Count : 0x%08X )" % (n_Position, t_Hdr.cste)
                return None, None
            dl_PropBagStore.append( l_stringTable )
            n_Position += n_Offset
            
            
        except :
            print format_exc()
            return None, None
        
        return dl_PropBagStore, n_Position

    def fnMappedFactoidType(self, s_SmartTagData, n_Position):
        
        try :
            
            MappedBasic = CMappedBasic()
            
            l_FactoidType = []
            
# Step 1-1. Get "FactoidType.cbFactoid" - FactoidType's Size ( exclusive itself ) 
            n_cbFactoid = CBuffer.fnReadDword(s_SmartTagData, n_Position)
            l_FactoidType.append( n_cbFactoid )
            n_Position += 4


# Step 1-2. Get "FactoidType.id" 
            n_id = CBuffer.fnReadDword(s_SmartTagData, n_Position)
            l_FactoidType.append( n_id )
            n_Position += 4


# Step 2. Mapped rgbUri / rgbTag / rgbDownloadURL ( All PBSting Type )
            l_FactoidType, n_Offset = MappedBasic.fnMappedPBStrings(s_SmartTagData[n_Position:], 3)
            if l_FactoidType == None or n_Offset == None :
                print "\t" * 4 + "[-] Failure - MappedPBStrings( Buffer : SmartTagData, Position : 0x%08X, Count : 3 )"
                return None, None
            n_Position += n_Offset

            
        except :
            print format_exc()
            return None
        
        return l_FactoidType, n_Position
    
    def fnMappedPropBags(self, s_PropBags):
        
        try :
            
            MappedBasic = CMappedBasic()
            
            dl_PropBags = []
            
# Step 1. Get "propBags.id, propBags.cProp, propBags.cbUnknown"
            n_Position = 0
            t_Hdr = MappedBasic.fnMappedPropBagHdr(s_PropBags[n_Position:n_Position + 0x6])
            if t_Hdr == None :
                print "\t" * 4 + "[-] Failure - MappedPropBagHdr()"
                return None
            dl_PropBags.append( t_Hdr )
            n_Position += 6
            
            
# Step 2. Mapped Properties
            l_Properties = self.fnMappedProperties(s_PropBags[n_Position:], t_Hdr.cProp)
            if l_Properties == None :
                print "\t" * 4 + "[-] Failure - MappedProperties()"
                return None
            dl_PropBags.append( l_Properties )
            
        except :
            print format_exc()
            return None
        
        return dl_PropBags

    def fnMappedProperties(self, s_Property, n_Count):

        try :
            
            l_PropPair = []
            
            n_Cnt = 0
            while True :
                if n_Cnt == n_Count :
                    break
                
                t_PropPair_Name = namedtuple("KeyValuePair", RULE_PROPBAG_PAIR_NAME)
                t_PropPair = t_PropPair_Name._make( unpack(RULE_PROPBAG_PAIR_PATTER, s_Property[n_Cnt * 8:n_Cnt * 8 + 8]) )
                l_PropPair.append( t_PropPair )
                n_Cnt += 1

            
            if s_Property.__len__() != n_Count * 8 : 
                print "\t" * 4 + "[-] Warning - Size UnMapped! Check plz.."
        
        except error :
            print "\t" * 4 + "[-] Error -  unpack requires a string argument of length 8 ( Length : 0x%08X )" % s_Property[n_Cnt * 8:n_Cnt * 8 + 8].__len__() 
            return None
        
        except :
            print format_exc()
            return None
        
        return l_PropPair

    def fnMappedSttbHdr(self, s_Sttb):
        
        try :
            
            d_SttbHdr = {}
            
# Step 1. Read fExtend
            n_Position = 0
            n_fExtend = CBuffer.fnReadWord(s_Sttb, n_Position)
            if n_fExtend != 0xFFFF :
                print "\t" * 4 + "[-] Error ( fExtend : %04X - MUST 0xFFFF )" % n_fExtend
                return None
            d_SttbHdr["fExtend"] = n_fExtend
            n_Position += 2
            
# Step 2. Read cData
            n_cData = CBuffer.fnReadWord(s_Sttb, n_Position)
            if n_cData > 0x7FF0 :
                print "\t" *  4 + "[-] Error ( cData : %04X - MUST NOT Exceed 0x7FF0 )" % n_cData
                return None
            d_SttbHdr["cData"] = n_cData
            n_Position += 2
            
# Step 3. Read cbExtra
            n_cbExtra = CBuffer.fnReadWord(s_Sttb, n_Position)
            if n_cbExtra != 0x0000 :
                print "\t" * 4 + "[-] Error ( cbExtra : %04X - MUST 0x0000 )" % n_cbExtra
                return None
            d_SttbHdr["cbExtra"] = n_cbExtra
            n_Position += 2
            
# Step 4. Read cchData0
            n_cchData0 = CBuffer.fnReadWord(s_Sttb, n_Position)
            if n_cchData0 != 0x0006 :
                print "\t" * 4 + "[-] Error ( cchData0 : %04X - MUST 0x0006 )" % n_cchData0
                return None
            d_SttbHdr["cchData0"] = n_cchData0
            
        except :
            print format_exc()
            return None
        
        return d_SttbHdr

    def fnMappedFbkfd(self, s_pBuf, l_StartOffset, s_Table, n_Offset, n_Size):
        
        try :
            
            Check = CCheck()
            MappedBasic = CMappedBasic()
            
            l_cp = []
            l_Fbkfd = []
            
            s_PlcfBkfd = s_Table[n_Offset : n_Offset + n_Size]
            if n_Offset + n_Size > s_PlcfBkfd.__len__() :
                print "\t" * 4 + "[*] CVE-2006-2492 : PlcfBkfFactoid is out of range! ( Offset : %08X, Size : %08X )" % (n_Offset, n_Size)
                
                # Enable Vulnerability
                s_PlcfBkfd = Check.fnCheckEnbleVulBuffer(s_pBuf, l_StartOffset, "1Table", n_Offset, n_Size)

            n_Position = 0
            n_Cnt = ( n_Size - 4 ) / 10
            while True :
                l_cp.append( CBuffer.fnReadDword(s_PlcfBkfd, n_Position))
                n_Position += 4
                
                l_Fbkfd.append( MappedBasic.fnMappedFBKFD( s_PlcfBkfd[n_Position:n_Position + SIZE_OF_FBKFD] ) )
                n_Position += SIZE_OF_FBKFD
                
                n_Cnt -= 1
                if n_Cnt == 0 :
                    break

            l_cp.append( CBuffer.fnReadDword(s_PlcfBkfd, n_Position))

        except :
            print format_exc()
            return None, None, None
        
        return l_cp, l_Fbkfd

    def fnMappedFbkld(self, s_pBuf, l_StartOffset, s_Table, n_Offset, n_Size):
        
        try :
            
            MappedBasic = CMappedBasic()
            
            l_cp = []
            l_Fbkld = []
            
            s_PlcfBkld = s_Table[n_Offset:n_Offset + n_Size]

            n_Position = 0
            n_Cnt = (n_Size - 4) / 8
            while True :
                l_cp.append( CBuffer.fnReadDword(s_PlcfBkld, n_Position))
                n_Position += 4
                
                l_Fbkld.append( MappedBasic.fnMappedFBKLD( s_PlcfBkld[n_Position:n_Position + SIZE_OF_FBKLD]) )
                n_Position += SIZE_OF_FBKLD
                
                n_Cnt -= 1
                if n_Cnt == 0 :
                    break
            
            l_cp.append( CBuffer.fnReadDword(s_PlcfBkld, n_Position))
            
        except:
            print format_exc()
            return None, None
        
        return l_cp, l_Fbkld


class CMappedBasic():
    
    def fnMappedPlc(self, s_Buffer, n_Offset, n_Size, n_End, n_ReadSize):
        
        try :
            
            n_Cnt = (n_Size / 4) / 2
            s_Plc = s_Buffer[n_Offset:n_Offset + n_Size]
            if s_Plc == "" :
                print "\t" * 4 + "[-] Error - Plc is Not in Table Stream. ( Table Buffer : 0x%08X, Offset: 0x%08X, Size : 0x%08X )" % (s_Buffer.__len__(), n_Offset, n_Size)
                return None, None
            
# Step 1. Get Plc.fc List            
            l_fc, s_PnFkp = self.fnMappedfc(s_Plc, n_Cnt + 1, n_End)
            if l_fc == None or s_PnFkp == None :
                print "\t" * 4 + "[-] Failure - Mappedfc()"
                return None, None
            
            
# Step 2. Get Plc.PnFkp.pn
            #    n_ReadSize <= 4    Return : Pn List
            #    n_ReadSize > 4     Return : Reading Data List
            l_Data = self.fnMappedPnFkp(s_PnFkp, n_Cnt, n_ReadSize)
            if l_Data == None :
                print "\t" * 4 + "[-] Failure - MappedPnFkp()"
                return None, None

        except :
            print format_exc()
            return None, None
        
        return l_fc, l_Data

    def fnMappedfc(self, s_Plc, n_Cnt, n_End):
        
        try :
            
            l_fc = []
            n_Position = 0
            
            if n_End != None :
                while True :
                    if CBuffer.fnReadByte(s_Plc, n_Position) == n_End :
                        break
                    n_fc = CBuffer.fnReadDword(s_Plc, n_Position)
                    n_Position += 4
                    l_fc.append( n_fc )
                    if n_Cnt == l_fc.__len__() :
                        break            
            else :
                while True :
                    n_fc = CBuffer.fnReadDword(s_Plc, n_Position)
                    n_Position += 4
                    l_fc.append( n_fc )
                    if n_Cnt == l_fc.__len__() :
                        break
            
            s_PnFkp = s_Plc[n_Position:]
            
        except :
            print format_exc()
            return None, None
        
        return l_fc, s_PnFkp
    
    def fnMappedPnFkp(self, s_PnFkp, n_Cnt, n_ReadSize):
        
        try :
            
            l_Data = []
            n_Position = 0
            
            while True :
                if n_ReadSize <= 4 :
                    n_tmpPnFkp = eval("CBuffer.fnRead" + g_Size[ n_ReadSize ])(s_PnFkp, n_Position)
                    n_PnFkp = self.fnMappedPn( n_tmpPnFkp )
                else :
                    n_PnFkp = s_PnFkp[n_Position:n_Position + n_ReadSize]
            
                n_Position += n_ReadSize
                l_Data.append( n_PnFkp )
                if n_Cnt == l_Data.__len__() :
                    break
            
        except :
            print format_exc()
            return None
        
        return l_Data
    
    def fnMappedPn(self, n_PnFkp):
        
        try :

            t_PnFkp_Name = namedtuple("PnFkpChpx", RULE_PNFKP_NAME)
            s_PnFkp = t_PnFkp_Name._make( CBuffer.fnBitUnpack(RULE_PNFKP_PATTERN, n_PnFkp) )

        except :
            print format_exc()
            return None
        
        return s_PnFkp.Pn

    def fnMappedCP(self, s_WordDoc, l_Pn, n_Max, n_Min, n_ReadSize):
        
        try :
            
            dl_Rg1 = []
            dl_Rg2 = []
            
            Check = CCheck()
            l_Cnt, l_Fkp = Check.fnCheckCP(s_WordDoc, l_Pn, n_Max, n_Min)
            if l_Cnt == None or l_Fkp == None :
                print "\t" * 4 + "[-] Failure - CheckCP()"
                return None, None
            if l_Cnt == [] or l_Fkp == [] :
                return l_Fkp, dl_Rg1, dl_Rg2
            
            for n_Index in range( l_Fkp.__len__() ) :
                l_Rg1 = []
                l_Rg2 = []
                s_Fkp = l_Fkp[n_Index]
                
                n_Offset = 0
                for n_Offset in range( l_Cnt[n_Index] + 1 ) :
                    l_Rg1.append( CBuffer.fnReadWord(s_Fkp, n_Offset * 4) )
                
                n_Position = (l_Cnt[n_Index] + 1) * 4
                
                if n_ReadSize <= 4 :
                    while True :
                        l_Rg2.append( eval("CBuffer.fnRead" + g_Size[ n_ReadSize ])(s_Fkp, n_Position) )
                        n_Position += n_ReadSize
                        if l_Cnt[n_Index] == l_Rg2.__len__() :
                            break
                        
                else :
                    while True :
                        l_Rg2.append( s_Fkp[n_Position:n_Position + n_ReadSize])
                        n_Position += n_ReadSize
                        if l_Cnt[n_Index] == l_Rg2.__len__() :
                            break
                
                dl_Rg1.append( l_Rg1 )
                dl_Rg2.append( l_Rg2 )
            

        except :
            print format_exc()
            return None, None
        
        return l_Fkp, dl_Rg1, dl_Rg2

    def fnMappedSprm(self, n_Sprm):
        
        try :
            
            t_Sprm_Name = namedtuple("Sprm", RULE_SPRM_NAME)
            t_Sprm = t_Sprm_Name._make( CBuffer.fnBitUnpack(RULE_SPRM_PATTERN, n_Sprm) )
            
        except :
            print format_exc()
            return None
        
        return t_Sprm
  
    def fnMappedPBString(self, s_Data):
        
        try :
            
            if s_Data == None :
                return None
            elif type(s_Data) != type(1) :
                print "\t" * 4 + "[-] Wrong Type Data. Need to Int ( Parameter Type : %s )" % type(s_Data)
                return None
            
            t_PBString_Name = namedtuple("PBString", RULE_PBSTRING_NAME)
            t_PBString = t_PBString_Name._make( CBuffer.fnBitUnpack(RULE_PBSTRING_PATTERN, s_Data) )

        except :
            print format_exc()
            return None
        
        return t_PBString
  
    def fnMappedPBStrings(self, s_Data, n_Count):
        
        try :
            
            l_PBStrings = []
            
            n_Cnt = 0
            n_Offset = 0
            while True :
                if n_Cnt == n_Count :
                    break
                
                l_tmp = []
                t_PBString = self.fnMappedPBString( CBuffer.fnReadWord(s_Data, n_Offset) )
                if t_PBString == None : 
                    print "\t" * 4 + "[-] Failure - MappedPBString()"
                    return None, None
                l_tmp.append( t_PBString )
                n_Offset += 2
                
                if t_PBString.fAnsiString == 1 :    # Type - Ansi
                    s_PBString = s_Data[n_Offset:n_Offset + t_PBString.cch]
                    n_Offset += t_PBString.cch
                else :  # t_PBString.fAnsiString == 0 : Type - Unicode 
                    s_PBString = s_Data[n_Offset:n_Offset + (t_PBString.cch * 2)]
                    n_Offset += t_PBString.cch * 2
                    
                l_tmp.append( s_PBString )
                
                l_PBStrings.append( l_tmp )
                
                n_Cnt += 1
  
        except :
            print format_exc()
            return None, None
  
        return l_PBStrings, n_Offset
       
    def fnMappedPropBagStoreHdr(self, s_Hdr):
        
        try :
            
            t_Hdr_Name = namedtuple("Hdr", RULE_PROPHDR1_NAME)
            t_Hdr = t_Hdr_Name._make( unpack(RULE_PROPHDR1_PATTERN, s_Hdr) )
            
            if t_Hdr.cbHdr != s_Hdr.__len__() or t_Hdr.cbHdr != 0xc :       # MUST Size : 0x0C
                print "\t" * 4 + "[-] Error ( Hdr Buffer Length : 0x%08X, Hdr.chHdr : 0x%08X )" % (s_Hdr.__len__(), t_Hdr.cbHdr)
                return None
            
            if t_Hdr.sVer != 0x0100 :           # MUST Version 0x0100
                print "\t" * 4 + "[-] Error ( Hdr Version : %04X )" % t_Hdr.sVer
                return None
            
        except :
            print format_exc()
            return None
        
        return t_Hdr
  
    def fnMappedPropBagHdr(self, s_Hdr):
        
        try :
            
            t_Hdr_Name = namedtuple("Hdr", RULE_PROPHDR2_NAME)
            t_Hdr = t_Hdr_Name._make( unpack(RULE_PROPHDR2_PATTERN, s_Hdr) )
            
            if t_Hdr.cbUnknown != 0x0000 :
                print "\t" * 4 + "[-] Warning ( Hdr cbUnknown : 0x%04X, MUST 0x0000 )" % t_Hdr.cbUnknown
            
        except :
            print format_exc()
            return None
        
        return t_Hdr
 
    def fnMappedFBKFD(self, s_FBKFD):
 
        try :
            
            l_FBKFD = []
            
            n_Position = 0
            l_FBKF = self.fnMappedFBKF( s_FBKFD[n_Position:n_Position+4])
            n_Position += 4
            
            cDepth = CBuffer.fnReadWord(s_FBKFD, n_Position)
            
            l_FBKFD.append( l_FBKF )
            l_FBKFD.append( cDepth )
            
        except :
            print format_exc()
            return None
        
        return l_FBKFD
 
    def fnMappedFBKF(self, s_FBKF):
        
        try :
            
            l_FBKF = []
            
            n_Position = 0
            n_ibkl = CBuffer.fnReadWord(s_FBKF, n_Position)
            n_Position += 2
            
            t_BKC = self.fnMappedBKC( CBuffer.fnReadWord(s_FBKF, n_Position) )
            
            l_FBKF.append( n_ibkl )
            l_FBKF.append( t_BKC )
            
        except :
            print format_exc()
            return None
        
        return l_FBKF
 
    def fnMappedBKC(self, n_BKC):
        
        try : 
            
            t_BKC_Name = namedtuple( "BKC", RULE_BKC_NAME )
            t_BKC = t_BKC_Name._make( CBuffer.fnBitUnpack(RULE_BKC_PATTERN, n_BKC))
            
            if t_BKC.fPub != 0 :
                print "\t" * 4 + "[-] Warning ( fPub : %d, MUST 0 )" % t_BKC.fPub
            
            if t_BKC.fCol != 0 :
                if t_BKC.itcFirst >= t_BKC.itcLim :
                    print "\t" * 4 + "[-] Error ( itcFirst : " + hex(t_BKC.itcFirst) + ", itcLim : " + hex(t_BKC.itcLim) + ")"
                    return None
            
        except :
            print format_exc()
            return None
        
        return t_BKC
 
    def fnMappedFBKLD(self, s_FBKLD):
        
        try :
            
            t_FBKLD_Name = namedtuple( "FBKLD", RULE_FBKLD_NAME )
            t_FBKLD = t_FBKLD_Name._make( unpack(RULE_FBKLD_PATTERN, s_FBKLD) )
            
        except :
            print format_exc()
            return None
        
        return t_FBKLD
 
    def fnMappedFACTOIDINFO(self, s_FactoidInfo):
    
        try :
            
            l_FactoidInfo = []
            n_Position = 0
            
# Step 1. Read dwId
            l_FactoidInfo.append( CBuffer.fnReadDword(s_FactoidInfo, n_Position) )
            n_Position += 4
            
# Step 2. Mapped SubEntity
            l_FactoidInfo.append( self.fnMappedSubEntity( CBuffer.fnReadWord(s_FactoidInfo, n_Position) ))
            n_Position += 2
            
# Step 3. Mapped FTO
            l_FactoidInfo.append( CBuffer.fnReadWord(s_FactoidInfo, n_Position) )
            n_Position += 2
            
# Step 4. Mapped pfpb
            l_FactoidInfo.append( CBuffer.fnReadDword(s_FactoidInfo, n_Position))
            
        except :
            print format_exc()
            return None

        return l_FactoidInfo

    def fnMappedSubEntity(self, n_SubEntity):
        
        try :
            
            t_SubEntity_Name = namedtuple( "SubEntity", RULE_SUBENTITY_NAME )
            t_SubEntity = t_SubEntity_Name._make( CBuffer.fnBitUnpack(RULE_SUBENTITY_PATTERN, n_SubEntity))
            
        except :
            print format_exc()
            return None
        
        return t_SubEntity


class CCheck():
    
    def fnCheckCP(self, s_WordDoc, l_Pn, n_Max, n_Min):
        
        try :
            
            l_Fkp = []
            l_Cnt = []
            
            for n_Pn in l_Pn :
                if n_Pn * 0x200 + 0x200 > s_WordDoc.__len__() :
                    print "\t" * 4 +"[-] Error - Out of Range! ( Pn : %04X, Offset : 0x%08X)" % (n_Pn, n_Pn * 0x200)
                    return None, None
                s_Fkp = s_WordDoc[n_Pn * 0x200 : n_Pn * 0x200 + 0x200]
                n_Cnt = CBuffer.fnConvertBinaryString2Int(s_Fkp[-1], 1)
                if n_Cnt <= n_Min -1 or n_Cnt > n_Max :
                    return l_Cnt, l_Fkp
                l_Cnt.append( n_Cnt )
                l_Fkp.append( s_Fkp )    
    
        except :
            print format_exc()
            return None, None
        
        return l_Cnt, l_Fkp
    
    def fnChckDuplicateList(self, l_Dup):
        
        try :
            
            l_NonDup = []
            
            for s_Dup in l_Dup :
                if s_Dup not in l_NonDup :
                    l_NonDup.append( s_Dup )
            
        except :
            print format_exc()
            return None
        
        return l_NonDup
    
    def fnCheckEnbleVulBuffer(self, s_pBuf, l_StartOffset, s_StreamName, n_Offset, n_Size):
        
        try :
            for l_Start in l_StartOffset :
                for n_Index in range( l_Start[0].__len__() ) :
                    if l_Start[0][n_Index] == s_StreamName :
                        n_Position = l_Start[1][n_Index]
                        return s_pBuf[n_Position+n_Offset:n_Position+n_Offset+n_Size]
                              
        except :
            print format_exc()
            return None


class CPrintDoc():
    
    def fnPlcBteChpx(self, s_WordDoc, s_Table, l_fc, l_Pn, l_Fkp, dl_Rgfc, dl_Rgb, dl_Group):
    
        try :
            
            s_OutPrint = ""
            s_OutPrint += "\t" * 1 + "- aFC\n"
            for n_Index in range( l_fc.__len__() ) :
                s_OutPrint += "\t" * 2 + "- FC[%04X] : 0x%08X\n" % (n_Index, l_fc[n_Index])
                
            s_OutPrint += "\t" * 1 + "- aPnBteChpx\n"
            for n_Index1 in range( l_Pn.__len__() ) :                
                s_OutPrint += "\t" * 2 + "- PnBteChpx.pn[%04X] : 0x%08X\n" % (n_Index1, l_Pn[n_Index1] * 0x200)
                
                for n_tmpIndex1 in range( dl_Rgfc.__len__() ) :
                    for n_tmpIndex2 in range( dl_Rgfc[n_tmpIndex1].__len__() ) :
                        s_OutPrint += "\t" * 3 + "- rgfc[%04X] : 0x%08X\n" % (n_tmpIndex2, dl_Rgfc[n_tmpIndex1][n_tmpIndex2])
                
                for n_Index2 in range( dl_Rgb.__len__() ) :
                    for n_Index3 in range( dl_Rgb[n_Index2].__len__() ) :
                        s_OutPrint += "\t" * 3 + "- rgb[%04X] : 0x%2X ( 0x%04X )\n" % (n_Index3, dl_Rgb[n_Index2][n_Index3], dl_Rgb[n_Index2][n_Index3] * 2)
                        s_OutPrint += "\t" * 4 + "- cb : 0x%02X\n" % (CBuffer.fnReadByte(l_Fkp[n_Index1], dl_Rgb[n_Index2][n_Index3] * 2))
                        
                        s_OutPrint += "\t" * 4 + "- GrpPrl\n"
                        for s_Group in dl_Group[n_Index2] :
                            n_Position = 0
                            while True :
                                n_Sprm = CBuffer.fnReadWord(s_Group, n_Position)
                                n_Position += 2
                                n_ispmd = n_Sprm & 0x01FF
                                n_f = (n_Sprm / 0x200) & 0x0001
                                n_sgc = (n_Sprm / 0x400) & 0x0007
                                n_spra = n_Sprm / 0x2000
                    
                                if n_spra == 0 :           n_szOperand = 1     
                                elif n_spra == 1 :         n_szOperand = 1     
                                elif n_spra == 2 :         n_szOperand = 2     
                                elif n_spra == 3 :         n_szOperand = 4     
                                elif n_spra == 4 :         n_szOperand = 2     
                                elif n_spra == 5 :         n_szOperand = 2     
                                elif n_spra == 6 :         n_szOperand = CBuffer.fnReadByte(s_Group, n_Position);  n_Position += 1
                                elif n_spra == 7 :         n_szOperand = 3
                                
                                s_Data = ""
                                n_Offset = n_Position
                                while n_Offset < n_Position + n_szOperand :
                                    s_Data += "0x%02X " % CBuffer.fnReadByte(s_Group, n_Offset)
                                    n_Offset += 1
                                
                                n_Position += n_szOperand
                                
                                s_OutPrint += "\t" * 5 + "- sprm : 0x%04X" % n_Sprm
                                n_NameIndex = eval(SINGLE_PROPERTY_MODIFIER[ n_sgc ].upper() + "_PROPERTY_MODIFIER_VALUE").index( n_Sprm )
                                if n_NameIndex != -1 :
                                    s_SPM_Name = eval(SINGLE_PROPERTY_MODIFIER[ n_sgc ].upper() + "_PROPERTY_MODIFIER_NAME")[n_NameIndex]
                                    s_OutPrint += " > %s\n" % s_SPM_Name
                                else :
                                    s_OutPrint += "\n"
                                s_OutPrint += "\t" * 6 + "- ispmd : " + bin(n_ispmd) + "\n"
                                s_OutPrint += "\t" * 6 + "- fSpec : " + bin(n_f) + "\n"
                                s_OutPrint += "\t" * 6 + "- sgc : " + bin(n_sgc) + "\n"
                                s_OutPrint += "\t" * 6 + "- spra : " + bin(n_spra) + "   (Size : %02X - %s)\n" % (n_szOperand, s_Data)
                                
                                if n_Position == s_Group.__len__() :
                                    break
                
                for n_Index4 in range( dl_Rgb.__len__() ) :
                    s_OutPrint += "\t" * 3 + "- crun : 0x%08X\n" % dl_Rgb[n_Index4].__len__()

        except :
            return format_exc()
        
        return s_OutPrint

    def fnPlcfBtePapx(self, s_WordDoc, s_Table, l_fc, l_Pn, l_Fkp, dl_Rgfc, dl_Rgbx, dl_istd, dl_Group):
        
        pass
    
    
    

class CVulnerability():

    # Embedded PE
    def fnGeneralPE(self, s_Buffer):
        
        try :

# Type 1. Unicode -> Ascii
            ConvertString = CConvertString()
            
            for s_CvtSig in g_ConvertSig :
                if s_Buffer[0:0x40].find(s_CvtSig) == -1 :
                    continue
                else :
                    n_Offset = s_Buffer[0:0x40].index(s_CvtSig)
            
                s_ConvertBuf = ConvertString.fnConvertStr2Hex(s_Buffer[n_Offset:], 2, None)
                if s_ConvertBuf == None :
                    print "\t" * 4 + "[-] Failure - ConvertStr2Hex()"
                    return False
                
                if s_ConvertBuf.find("MZ") or s_ConvertBuf.find("mz") :
                    try :
                        n_Index = s_ConvertBuf.index("MZ")
                    except :
                        n_Index = s_ConvertBuf.index("mz")
                        
                    s_Format = CPE.fnCheck(s_ConvertBuf[n_Index:])
                    if s_Format == "PE" :
                        print "\t" * 4 + "[*] Possible - Found PE Files!"
                        break
            
        except :
            print format_exc()
            return False
        
        return True
    
    def fnCVE200062492(self, l_FactoidInfo):
        
        try :
            
            if l_FactoidInfo[3] != 0 :
                print "\t" * 4 + "[*] CVE-2006-2492 : SmartTag ( Pointer : 0x%08X )" % l_FactoidInfo[3]
            
        except :
            print format_exc()
            return None
        
        return True


