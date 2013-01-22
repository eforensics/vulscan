

#######################################
#    Struct.Unpack( ) Format
#######################################
#1Byte
#    Char                    c
#    Signed Char             b
#    Unsigned Char           B
#    BOOL                    ?
#2Bytes ( Little-Endian )
#    Short                   h
#    Unsigned Short          H
#4Bytes ( Little-Endian )
#    Int                     i
#    Unsigned Int            I
#    Long                    l
#    Unsigned Long           L
#    Float                   f
#8Bytes ( Little-Endian )
#    Long Long               q
#    Unsigned Long Long      Q
#    Double                  d
#String
#    Char[]                  s, p
#Point
#    Void *                  P    
#######################################


# OLE-Common
SIZE_OF_SECTOR = 0x200
SIZE_OF_SHORT_SECTOR = 0x40

SIZE_OF_OLE_HEADER = 0x200
RULE_OLEHEADER_PATTERN = '=8s16s5h10s8l436s'
RULE_OLEHEADER_NAME = ('Signature UID Revision Version ByteOrder szSector szShort Reserved1 NumSAT DirSecID Reserved2 szMinSector SSATSecID NumSSAT MSATSecID NumMSAT MSAT')

SIZE_OF_DIRECTORY = 0x80
RULE_OLEDIRECTORY_PATTERN = '=64sH2B3l16sl8s8s3l'
RULE_OLEDIRECTORY_NAME = ('DirName Size Type Color LeftChild RightChild RootChild UID Flags createTime LastModifyTime SecID_Start StreamSize Reserved')



# OLE-HWP
RULE_HWPHEADER_PATTERN = '=32s2l216s'
RULE_HWPHEADER_NAME = ('Signature Version Property Reserved')

          #    TagID 
HWPTAG_BEGIN = 0x010
          #    DocInfo
HWPTAG = {HWPTAG_BEGIN      :"HWPTAG_DOCUMENT_PROPERTIES",
          HWPTAG_BEGIN+1    :"HWPTAG_ID_MAPPINGS",
          HWPTAG_BEGIN+2    :"HWPTAG_BIN_DATA",
          HWPTAG_BEGIN+3    :"HWPTAG_FACE_NAME",
          HWPTAG_BEGIN+4    :"HWPTAG_BORDER_FILL",
          HWPTAG_BEGIN+5    :"HWPTAG_CHAR_SHAPE",
          HWPTAG_BEGIN+6    :"HWPTAG_TAB_DEF",
          HWPTAG_BEGIN+7    :"HWPTAG_NUMBERING",
          HWPTAG_BEGIN+8    :"HWPTAG_BULLET",
          HWPTAG_BEGIN+9    :"HWPTAG_PARA_SHAPE",
          HWPTAG_BEGIN+10   :"HWPTAG_STYLE",
          HWPTAG_BEGIN+11   :"HWPTAG_DOC_DATA",
          HWPTAG_BEGIN+12   :"HWPTAG_DISTRIBUTE_DOC_DATA",
          HWPTAG_BEGIN+13   :"RESERVED",
          HWPTAG_BEGIN+14   :"HWPTAG_COMPATIBLE_DOCUMENT",
          HWPTAG_BEGIN+15   :"HWPTAG_LAYOUT_COMPATIBILITY",
          
          # BodyText 
          HWPTAG_BEGIN+50   :"HWPTAG_PARA_HEADER",
          HWPTAG_BEGIN+51   :"HWPTAG_PARA_TEXT",
          HWPTAG_BEGIN+52   :"HWPTAG_PARA_CHAR",
          HWPTAG_BEGIN+53   :"HWPTAG_PARA_LINE_SEG",
          HWPTAG_BEGIN+54   :"HWPTAG_PARA_RANGE_TAG",
          HWPTAG_BEGIN+55   :"HWPTAG_CTRL_HEADER",
          HWPTAG_BEGIN+56   :"HWPTAG_LIST_HEADER",
          HWPTAG_BEGIN+57   :"HWPTAG_PAGE_DEF",
          HWPTAG_BEGIN+58   :"HWPTAG_FOOTNOTE_SHAPE",
          HWPTAG_BEGIN+59   :"HWPTAG_PAGE_BORDER_FILL",
          HWPTAG_BEGIN+60   :"HWPTAG_SHAPE_COMPONENT",
          HWPTAG_BEGIN+61   :"HWPTAG_TABLE",
          HWPTAG_BEGIN+62   :"HWPTAG_SHAPE_COMPONENT_LINE",
          HWPTAG_BEGIN+63   :"HWPTAG_SHAPE_COMPONENT_RECTANGLE",
          HWPTAG_BEGIN+64   :"HWPTAG_SHAPE_COMPONENT_ELLIPSE",
          HWPTAG_BEGIN+65   :"HWPTAG_SHAPE_COMPONENT_ARC",
          HWPTAG_BEGIN+66   :"HWPTAG_SHAPE_COMPONENT_POLYGON",
          HWPTAG_BEGIN+67   :"HWPTAG_SHAPE_COMPONENT_CURVE",
          HWPTAG_BEGIN+68   :"HWPTAG_SHAPE_COMPONENT_OLE",
          HWPTAG_BEGIN+69   :"HWPTAG_SHAPE_COMPONENT_PICTURE",
          HWPTAG_BEGIN+70   :"HWPTAG_SHAPE_COMPONENT_CONTAINER",
          HWPTAG_BEGIN+71   :"HWPTAG_CTRL_DATA",
          HWPTAG_BEGIN+72   :"HWPTAG_EQEDIT",
          HWPTAG_BEGIN+73   :"RESERVED",
          HWPTAG_BEGIN+74   :"HWPTAG_SHAPE_COMPONENT_TEXTART",
          HWPTAG_BEGIN+75   :"HWPTAG_FORM_OBJECT",
          HWPTAG_BEGIN+76   :"HWPTAG_SHAPE",
          HWPTAG_BEGIN+77   :"HWPTAG_MEMO_LIST",
          
          # DocInfo          
          HWPTAG_BEGIN+78   :"HWPTAG_FORBIDDEN_CHAR",
          
          # BodyText
          HWPTAG_BEGIN+79   :"HWPTAG_CHART_DATA",
          HWPTAG_BEGIN+99   :"HWPTAG_SHAPE_COMPONENT_UNKNOWN" }




# OLE-Office
g_Version = ["97", "2000", "2002", "2003", "2007"]
g_VersionCount = [186, 30, 56, 56, 38]

SIZE_OF_FIBBASE = 0x20      # 32Bytes
RULE_FIBBASE_PATTERN = '=7H1L2B2H2L'
RULE_FIBBASE_NAME = ('wIdent nFib unused lid pnNext Flag nFibBack lKey envr envtFlag reserved3 reserved4 reserved5 reserved6')


SIZE_OF_FIBW97 = 0x1C        # 28Bytes
RULE_FIBW97_PATTERN = '=14H'
RULE_FIBW97_NAME = ('reserved1 reserved2 reserved3 reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 lidFE')


SIZE_OF_FIBLW97 = 0x58      # 88Bytes
RULE_FIBLW97_PATTERN = '=22L'
RULE_FIBLW97_NAME = ('cbMac reserved1 reserved2 ccpText ccpFtn ccpHdd reserved3 ccpAtn ccpEdn ccpTxbx ccpHdrTxbx reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 reserved14')


g_Fib = [0x00C1, 0x00D9, 0x0101, 0x010C, 0x0112]
g_cbRgFcLcb = [0x005D, 0x006C, 0x0088, 0x00A4, 0x00B7]
g_cswNew = [0x0000, 0x0002, 0x0002, 0x0002, 0x0005]


g_StructList = ["PlcBteChpx"]
g_ExceptList = ["dwLowDateTime", "dwHighDateTime"]


SIZE_OF_FIBFCLCB_97 = 0x2E8       # 744Bytes
RULE_FIBFCLCB_PATTERN_97 = '=186L'
RULE_FIBFCLCB_NAME_97 = ('fcStshfOrig lcbStshfOrig fcStshf lcbStshf fcPlcffndRef lcbPlcffndRef fcPlcffndTxt lcbPlcffndTxt fcPlcfandRef lcbPlcfandRef fcPlcfandTxt lcbPlcfandTxt fcPlcfSed lcbPlcfSed fcPlcPad lcbPlcPad fcPlcfPhe lcbPlcfPhe fcSttbfGlsy lcbSttbfGlsy fcPlcfGlsy lcbPlcfGlsy fcPlcfHdd lcbPlcfHdd fcPlcBteChpx lcbPlcBteChpx fcPlcfBtePapx lcbPlcfBtePapx fcPlcfSea lcbPlcfSea fcSttbfFfn lcbSttbfFfn fcPlcfFldMom lcbPlcfFldMom fcPlcfFldHdr lcbPlcfFldHdr fcPlcfFldFtn lcbPlcfFldFtn fcPlcfFldAtn lcbPlcfFldAtn fcPlcfFldMcr lcbPlcfFldMcr fcSttbfBkmk lcbSttbfBkmk fcPlcfBkf lcbPlcfBkf fcPlcfBkl lcbPlcfBkl fcCmds lcbCmds fcUnused1 lcbUnused1 fcSttbfMcr lcbSttbfMcr fcPrDrvr lcbPDrvr fcPrEnvPort lcbPrEnvPort fcPrEnvLand lcbPrEnvLand fcWss lcbWss fcDop lcbDop fcSttbfAssoc lcbSttbfAssoc fcClx lcbClx fcPlcfPgdFtn lcbPlcfPgdFtn fcAutosaveSource lcbAutosaveSource fcGrpXstAtnOwners lcbGrpXstAtnOwners fcSttbfAtnBkmk lcbSttbfAtnBkmk fcUnused3 lcbUnused3 fcUnused2 lcbUnused2 fcPlcSpaMon lcbPlcSpaMon fcPlcSpaHdr lcbPlcSpaHdr fcPlcAtnBkf lcbPlcAtnBkf fcPlcAtnBkl lcbPlcAtnBkl fcPms lcbPms fcFormFldSttbs lcbFormFldSttbs fcPlcfendRef lcbPlcfendRef fcPlcfendTxt lcbPlcfendTxt fcPlcfFldEdn lcbPlcfFldEdn fcUnused4 lcbUnused4 fcDggInfo lcbDggInfo fcSttbfRMark lcbSttbfRMark fcSttbfCaption lcbSttbfCaption fcSttbfAutoCaption lcbSttbfAutoCaption fcPlcWkb lcbPlcWkb fcPlcfSpl lcbPlcfSpl fcPlcftxbxTxt lcbPlcftxbxTxt fcPlcfFldTxbx lcbPlcfFldTxbx fcPlcfHdrtxbxTxt lcbPlcfHdrtxbxTxt fcPlcffldHdrTxdx lcbPlcffldHdrTxdx fcStwUser lcbStwUser fcSttbTtmbd lcbSttbTtmbd fcCookieData lcbCookieData fcPgdMotherOldOld lcbPgdMotherOldOld fcBkdMotherOldOld lcbBkdMotherOldOld fcPgdFtnOldOld lcbPgdFtnOldOld fcBkdFtnOldOld lcbBkdFtnOldOld fcPgdEdnOldOld lcbPgdEdnOldOld fcBkdEdnOldOld lcbBkdEdnOldOld fcSttbfIntlFld lcbSttbfIntlFld fcRouteSlip lcbRouteSlip fcSttbSavedBy lcbSttbSavedBy fcSttbFnm lcbSttbFnm fcPlfLst lcbPlfLst fcPlfLfo lcbPlfLfo fcPlcTxbxBkd lcbPlcTxbxBkd fcPlcfTxbxHdrBkd lcbPlcfTxbxHdrBkd fcDocUndoWord9 lcbDocUndoWord9 fcRgbUse lcbRgbUse fcUsp lcbUsp fcUskf lcbUskf fcPlcupcRgbUse lcbPlcupcRgbUse fcPlcupcUsp lcbPlcupcUsp fcSttbGlsyStyle lcbSttbGlsyStyle fcPlgosl lcbPlgosl fcPlcocx lcbPlcocx fcPlcfBteLvc lcbPlcfBteLvc dwLowDateTime dwHighDateTime fcPlcfLvcPe10 lcbPlcfLvcPe10 fcPlcfAsumy lcbPlcfAsumy fcPlcfGram lcbPlcfGram fcSttbListNames lcbSttbListNames fcSttbfUssr lcbSttbfUssr')


SIZE_OF_FIBFCLCB_2000 = 0x78      # 120Bytes
RULE_FIBFCLCB_PATTERN_2000 = '=30L'
RULE_FIBFCLCB_NAME_2000 = ('fcPlcfTch lcbPlcfTch fcRmdThreading lcbRmdThreading fcMid lcbMid fcSttbRgtplc lcbSttbRgtplc fcMsoEnvelope lcbMsoEnvelope fcPlcfLad lcbPlcfLad fcRgDofr lcbRgDofr fcPlcosl lcbPlcosl fcPlcfCookieOld lcbPlcfCookieOld fcPgdMotherOld lcbPgdMotherOld fcBkdMotherOld lcbBkdMotherOld fcPgdFtnOld lcbPgdFtnOld fcBkdFtnOld lcbBkdFtnOld fcPgdEdnOld lcbPgdEdnOld fcBkdEdnOld lcbBkdEdnOld')


SIZE_OF_FIBFCLCB_2002 = 0xE0      # 224Bytes
RULE_FIBFCLCB_PATTERN_2002 = '=56L'
RULE_FIBFCLCB_NAME_2002 = ('fcUnused1 lcbUnused1 fcPlcfPgp lcbPlcfPgp fcPlcfuim lcbPlcfuim fcPlfguidUim lcbPlfguidUim fcAtrdExtra lcbAtrdExtra fcPlrsid lcbPlrsid fcSttbfBkmkFactoid lcbSttbfBkmkFactoid fcPlcfBkfFactoid lcbPlcfBkfFactoid fcPlcfcookie lcbPlcfcookie fcPlcfBklFactoid lcbPlcfBklFactoid fcFactoidData lcbFactoidData fcDocUndo lcbDocUndo fcSttbfBkmkFcc lcbSttbfBkmkFcc fcPlcfBkfFcc lcbPlcfBkfFcc fcPlcfBklFcc lcbPlcfBklFcc fcSttbfbkmkBPRepairs lcbSttbfbkmkBPRepairs fcPlcfbkfBPRepairs lcbPlcfbkfBPRepairs fcPlcfbklBPRepairs lcbPlcfbklBPRepairs fcPmsNew lcbPmsNew fcODSO lcbODSO fcPlcfpmiOldXP lcbPlcfpmiOldXP fcPlcfpmiNewXP lcbPlcfpmiNewXP fcPlcfpmiMixedXP lcbPlcfpmiMixedXP fcUnused2 lcbUnused2 fcPlcffactoid lcbPlcffactoid fcPlcflvcOldXP lcbPlcflvcOldXP fcPlcflvcNewXP lcbPlcflvcNewXP fcPlcflvcMixedXP lcbPlcflvcMixedXP')


SIZE_OF_FIBFCLCB_2003 = 0xE0      # 224Bytes
RULE_FIBFCLCB_PATTERN_2003 = '=56L'
RULE_FIBFCLCB_NAME_2003 = ('fcHplxsdr lcbHplxsdr fcSttbfBkmkSdt lcbSttbfBkmkSdt fcPlcfBkfSdt lcbPlcfBkfSdt fcPlcfBklSdt lcbPlcfBklSdt fcCustomXForm lcbCustomXForm fcSttbfBkmkProt lcbSttbfBkmkProt fcPlcfBkfProt lcbPlcfBkfProt fcPlcfBklProt lcbPlcfBklProt fcSttbProtUser lcbSttbProtUser fcUnused lcbUnused fcPlcfpmiOld lcbPlcfpmiOld fcPlcfpmiOldInline lcbPlcfpmiOldInline fcPlcfpmiNew lcbPlcfpmiNew fcPlcfpmiNewInline lcbPlcfpmiNewInline fcPlcflvcOld lcbPlcflvcOld fcPlcflvcOldInline lcbPlcflvcOldInline fcPlcflvcNew lcbPlcflvcNew fcPlcflcxNewInline lcbPlcflcxNewInline fcPgdMother lcbPgdMother fcBkdMother lcbBkdMother fcAfdMother lcbAfdMother fcPgdFtn lcbPgdFtn fcBkdFtn lcbBkdFtn fcAfdFtn lcbAfdFtn fcPgdEdn lcbPgdEdn fcBkdEdn lcbBkdEdn fcAfdEdn lcbAfdEdn fcAfd lcbAfd')


SIZE_OF_FIBFCLCB_2007 = 0x98      # 152Bytes
RULE_FIBFCLCB_PATTERN_2007 = '=38L'
RULE_FIBFCLCB_NAME_2007 = ('fcPlcfmthd lcbPlcfmthd fcSttbfBkmkMoveFrom lcbSttbfBkmkMoveFrom fcPlcfBkfMoveFrom lcbPlcfBkfMoveFrom fcPlcfBklMoveFrom lcbPlcfBklMoveFrom fcSttbfBkmkMoveTo lcbSttbfBkmkMoveTo fcPlcfBkfMoveTo lcbPlcfBkfMoveTo fcPlcfBklMoveTo lcbPlcfBklMoveTo fcUnused1 lcbUnused1 fcUnused2 lcbUnused2 fcUnused3 lcbUnused3 fcSttbfBkmkArto lcbSttbfBkmkArto fcPlcfBkfArto lcbPlcfBkfArto fcPlcfBklArto lcbPlcfBklArto fcArtoData lcbArtoData fcUnused4 lcbUnused4 fcUnused5 lcbUnused5 fcUnused6 lcbUnused6 fcOssTheme lcbOssTheme fcColorSchemeMapping lcbColorSchemeMapping')



# OLE-Office : Sub-Structure
RULE_PNFKPCHPX_PATTERN = '22,10'
RULE_PNFKPCHPX_NAME = ('Pn Unused') 


RULE_SPRM_PATTERN = '9,1,3,3'
RULE_SPRM_NAME = ('ispmd A sgc spra')


# OLE-Office : Single Property Modifier
SINGLE_PROPERTY_MODIFIER = {1:"Para", 2:"Char", 3:"Pict", 4:"Sect", 5:"Table"}

CHAR_PROPERTY_MODIFIER_VALUE = [0x0800, 0x0801, 0x0802, 0x6A03, 0x4804, 0x6805, 0x0806, 0x4807, 0x6A09, 0x080A, 0x2A0C, 0x0811, 0x6815, 0x6816, 0x6817, 0x0818, 0xC81A, 0x4A30, 0xCA31, 0x2A33, 0x2A34, 0x0835, 0x0836, 0x0838, 0x0838, 0x0839, 0x083A, 0x083B, 0x083C, 0x2A3E, 0x8840, 0x2A42, 0x4A43, 0x4845, 0xCA47, 0x2A48, 0x484B, 0x484E, 0x4A4F, 0x4A50, 0x4A51, 0x4852, 0x2A53, 0x0854, 0x0855, 0x0856, 0xCA57, 0x0858, 0x2859, 0x085A, 0x085C, 0x085D, 0x4A5E, 0x485F, 0x4A60, 0x4A61, 0xCA62, 0x4863, 0x6864, 0x6865, 0x4866, 0x4867, 0x0868, 0x486D, 0x486E, 0x286F, 0x6870, 0xCA71, 0xCA72, 0x4873, 0x4874, 0x0875, 0xCA76, 0x6877, 0xCA78, 0x2879, 0x0882, 0x2A83, 0xCA85, 0x2A86, 0x6887, 0x4888, 0xCA89, 0x2A90]
CHAR_PROPERTY_MODIFIER_NAME = ["sprmCFRMarkDel", "sprmCFRMarkIns", "sprmCFFldVanish", "sprmCPicLocation", "sprmCIbstRMark", "sprmCDttmRMark", "sprmCFData", "sprmCIdslRMark", "sprmCSymbol", "sprmCFOle2", "sprmCHighlight", "sprmCFWebHidden", "sprmCRsidProp", "sprmCRsidText", "sprmCRsidRMDel", "sprmCFSpecVanish", "sprmCFMathPr", "sprmCIstd", "sprmCIstdPermute", "sprmCPlain", "sprmCKcd", "sprmCFBold", "sprmCFItalic", "sprmCFStrike", "sprmCFOutline", "sprmCFShadow", "sprmCFSmallCaps", "sprmCFCaps", "sprmCFVanish", "sprmCKul", "sprmCDxzSpace", "sprmCIco", "sprmCHps", "sprmCHpsPos", "sprmCMajority", "sprmCIss", "sprmCHpsKern", "sprmCHresi", "sprmCRgFtc0", "sprmCRgFtc1", "sprmCRgFtc2", "sprmCCharScale", "sprmCFDStrike", "sprmCFImprint", "sprmCFSpec", "sprmCFObj", "sprmCPropRMark90", "sprmCFEmboss", "sprmCSfxText", "sprmCFBiDi", "sprmCFBoldBi", "sprmCFItalicBi", "sprmCFtcBi", "sprmClidBi", "sprmCIcoBi", "sprmCHpsBi", "sprmCDispFldRMark", "sprmCIbstRMarkDel", "sprmCDttmRMarkDel", "sprmCBrc80", "sprmCShd80", "sprmCIdslRMarkDel", "sprmCFUsePgsuSettings", "sprmCRgLid0_80", "sprmRgLid1_80", "sprmCIdctHint", "sprmCCv", "sprmCShd", "sprmCBrc", "sprmCRgLid0", "sprmCRgLid1", "sprmCFNoProof", "sprmCFitText", "sprmCCvUl", "sprmCFELayout", "sprmCLbcCRJ", "sprmCFComplexScripts", "sprmCWall", "sprmCCnf", "sprmCNeedFontFixup", "sprmCPbiIBullet", "sprmCPbiGrf", "sprmCPropRMark", "sprmCFSdtVanish"]

PARA_PROPERTY_MODIFIER_VALUE = [0x4600, 0xC601, 0x2602, 0x2403, 0x2405, 0x2406, 0x2407, 0x260A, 0x460B, 0x240C, 0xC60D, 0x840E, 0x840F, 0x4610, 0x8411, 0x6412, 0xA413, 0xA414, 0xC615, 0x2416, 0x2417, 0x8418, 0x8419, 0x841A, 0x261B, 0x2423, 0x6424, 0x6425, 0x6426, 0x6427, 0x6428, 0x6629, 0x242A, 0x442B, 0x442C, 0x442D, 0x842E, 0x842F, 0x2430, 0x2431, 0x2433, 0x2434, 0x2435, 0x2436, 0x2437, 0x2438, 0x4439, 0x443A, 0x2640, 0x2441, 0x2443, 0xC645, 0x6646, 0x2447, 0x2448, 0x6649, 0x664A, 0x244B, 0x244C, 0xC64D, 0xC64E, 0xC64F, 0xC650, 0xC651, 0xC652, 0xC653, 0x4455, 0x4456, 0x4457, 0x4458, 0x4459, 0x245A, 0x245B, 0x245C, 0x845D, 0x845E, 0x465F, 0x8460, 0x2461, 0x2462, 0x2664, 0x6465, 0xC666, 0x6467, 0xC669, 0x646B, 0xC66C, 0x246D, 0xC66F, 0x2470, 0x2471]
PARA_PROPERTY_MODIFIER_NAME = ["sprmPIstd", "sprmPIstdPermute", "sprmPIncLvl", "sprmPJc80", "sprmPFKeep", "sprmPFKeepFollow", "sprmPFPageBreakBefore", "sprmPIlvl", "sprmPIlfo", "sprmPFNoLineNumb", "sprmPChgTabsPapx", "sprmPDxaRight80", "sprmPDxaLeft80", "sprmPNest80", "sprmPDxaLeft180", "sprmPDyaLine", "sprmPDyaBefore", "sprmPDyaAfter", "sprmPChgTabs", "sprmPFInTable", "sprmPFTtp", "sprmPDxaAbs", "sprmPDyaAbs", "sprmPDxaWidth", "sprmPPc", "sprmPWr", "sprmPBrcTop80", "sprmPBrcLeft80", "sprmPBrcBottom80", "sprmPBrcRight80", "sprmPBrcBetween80", "sprmPBrcBar80", "sprmPFNoAutoHyph", "sprmPWHeightAbs", "sprmPDcs", "sprmPShd80", "sprmPDyaFromText", "sprmPDxaFromText", "sprmPFLocked", "sprmPFWidowControl", "sprmPFKinsoku", "sprmPFWordWrap", "sprmPDOverflowPunct", "sprmPFTopLinePunct", "sprmPFAutoSpaceDE", "sprmPFAutoSpaceDN", "sprmPWAlignFont", "sprmPFrameTextFlow", "sprmPOutLvl", "sprmPFBiDi", "sprmPFNumRMIns", "sprmPNumRM", "sprmPHugePapx", "sprmPFUsePgsuSettings", "sprmPFAdjustRight", "sprmPItap", "sprmPDtap", "sprmPFInnerTableCell", "sprmPFInnerTtp", "sprmPShd", "sprmPBrcTop", "sprmPBrcLeft", "sprmPBrcBottom", "sprmPBrcRight", "sprmPBrcBetween", "sprmPBrcBar", "sprmPDxcRight", "sprmPDxcLeft", "sprmPDxcLeft1", "sprmPDylBefore", "sprmPDylAfter", "sprmPFOpenTch", "sprmPFDyaBeforeAuto", "sprmPFDyaAfterAuto", "sprmPDxaRight", "sprmPDxaLeft", "sprmPNest", "sprmPDxaLeft1", "sprmPJc", "sprmPFNoAllowOverlap", "sprmPWall", "sprmPIpgp", "sprmPCnf", "sprmPRsid", "sprmPIstdListPermute", "sprmPTableProps", "sprmPTIstdInfo", "sprmPFContextualSpacing", "sprmPPropRMark", "sprmPFMirrorIndents", "sprmPTtwo"]

TABLE_PROPERTY_MODIFIER_VALUE = [0x5400, 0x9601, 0x9602, 0x3403, 0x3404, 0xD605, 0x9407, 0xD608, 0xD609, 0x740A, 0x560B, 0xD60C, 0x360D, 0x940E, 0x940F, 0x9410, 0x9411, 0xD612, 0xD613, 0xF614, 0x3615, 0xD616, 0xF617, 0xF618, 0x3619, 0xD61A, 0xD61B, 0xD61C, 0xD61D, 0x941E, 0x941F, 0xD620, 0x7621, 0x5622, 0x7623, 0x5624, 0x5625, 0x7629, 0xD62B, 0xD62C, 0xD62D, 0xD62E, 0xD62F, 0xD632, 0xD633, 0xD634, 0xD635, 0xF636, 0xD639, 0x563A, 0xD63E, 0xD642, 0xD660, 0xF661, 0xD662, 0x5664, 0x3465, 0x3466, 0xD667, 0x3668, 0x7469, 0xD66A, 0xD670, 0xD671, 0xD672, 0x7479, 0x347C, 0x347D, 0xD47F, 0xD680, 0xD681, 0xD682, 0xD683, 0xD684, 0xD685, 0xD686, 0xD687, 0x3488, 0x3489, 0x548A]
TABLE_PROPERTY_MODIFIER_NAME = ["sprmTJc90", "sprmTDxaLeft", "sprmTDxaGapHalf", "sprmTFCantSplit90", "sprmTTableHeader", "sprmTTableBorders80", "sprmTDyaRowHeight", "sprmTDefTable", "sprmTDefTableShd80", "sprmTTlp", "sprmTFBiDi", "sprmTDefTableShd3rd", "sprmTPc", "sprmTDxaAbs", "sprmTDyaAbs", "sprmTDxaFromText", "sprmTDyaFromText", "sprmTDefTableShd", "sprmTTableBorders", "sprmTTableWidth", "sprmTFAutofit", "sprmTDefTableShd2nd", "sprmTWidthBefore", "sprmTWidthAfter", "sprmTFKeepFollow", "sprmTBrcTopCv", "sprmTBrcLeftCv", "sprmTBrcBottomCv", "sprmTBrcRightCv", "sprmTDxaFromTextRight", "sprmTDyaFromTextBottom", "sprmTSetBrc80", "sprmTInsert", "sprmTDelete", "sprmTDxaCol", "sprmTMerge", "sprmTSplit", "sprmTTextFlow", "sprmTVertMerge", "sprmTVertAlign", "sprmTSetShd", "sprmTSetShdOdd", "sprmTSetBrc", "sprmTCellPadding", "sprmTCellSpacingDefault", "sprmTCellPaddingDefault", "sprmTCellWidth", "sprmTFitText", "sprmTFCellNoWrap", "sprmTIstd", "sprmTCellPaddingStyle", "sprmTCellFHideMark", "sprmTSetShdTable", "sprmTWidthIndent", "sprmTCellBrcType", "sprmTFBiDi90", "sprmTFNoAllowOverlap", "sprmTFCantSplit", "sprmTPropRMark", "sprmTWall", "sprmTIpgp", "sprmTCnf", "sprmTDefTableShdRaw", "sprmTDefTableShdRaw2nd", "sprmTDefTableShdRaw3rd", "sprmTRsid", "sprmTCellVertAlignStyle", "sprmTCellNoWrapStyle", "sprmTCellBrcTopStyle", "sprmTCellBrcBottomStyle", "sprmTCellBrcLeftStyle", "sprmTCellBrcRightStyle", "sprmTCellBrcInsideHStyle", "sprmTCellBrcInsideVStyle", "sprmTCellBrcTL2BRStyle", "sprmTCellBrcTR2BLStyle", "sprmTCellShdStyle", "sprmTCHorzBands", "sprmTCVertBands", "sprmTJc"]

SECT_PROPERTY_MODIFIER_VALUE = [0x3000, 0x3001, 0xF203, 0xF204, 0x3005, 0x3006, 0x5007, 0x5008, 0x3009, 0x300A, 0x500B, 0x900C, 0x300E, 0x3011, 0x3012, 0x3013, 0x5015, 0x9016, 0xB017, 0xB018, 0x3019, 0x301A, 0x501B, 0x501C, 0x301D, 0xB01F, 0xB020, 0xB021, 0xB022, 0x9023, 0x9024, 0xB025, 0x5026, 0x3228, 0x322A, 0x702B, 0x702C, 0x702D, 0x702E, 0x522F, 0x7030, 0x9031, 0x5032, 0x5033, 0xD234, 0xD235, 0xD236, 0xD237, 0x3239, 0x703A, 0x303B, 0x303C, 0x303E, 0x503F, 0x5040, 0x5041, 0x5042, 0xD243, 0x7044]
SECT_PROPERTY_MODIFIER_NAME = ["sprmScnsPgn", "sprmSiHeadingPgn", "sprmSDxaColWidth", "sprmSDxaColSpacing", "sprmSFEvenlySpaced", "sprmSFProtected", "sprmSDmBinFirst", "sprmSDmBinOther", "sprmSDkc", "sprmSFTitlePage", "sprmSCcolumns", "sprmSDxaColumns", "sprmSNfcPgn", "sprmSFPgnRestart", "sprmSFEndnote", "sprmSLnc", "sprmSNLnnMod", "sprmSDxaLnn", "sprmSDyaHdrTop", "sprmSDyaHdrBottom", "sprmSLBetween", "sprmSVjc", "sprmSLnnMin", "sprmSPgnStart97", "sprmSBOrientation", "sprmSXaPage", "sprmSYaPage", "sprmSDxaLeft", "sprmSDxaRight", "sprmSDyaTop", "sprmSDyaBottom", "sprmSDzaGutter", "sprmSDmPaperReq", "sprmSFBiDi", "sprmSFRTLGutter", "sprmSBrcTop80", "sprmSBrcLeft80", "sprmSBrcBottom80", "sprmSBrcRight80", "sprmSPgbProp", "sprmSDxtCharSpace", "sprmSDyaLinePitch", "sprmSClm", "sprmSTextFlow", "sprmSBrcTop", "sprmSBrcLeft", "sprmSBrcBottom", "sprmSBrcRight", "sprmSWall", "sprmSRsid", "sprmSFpc", "sprmSRncFtn", "sprmSRncEdn", "sprmSNFtn", "sprmSNfcFtnRef", "sprmSNEdn", "sprmSNfcEdnRef", "sprmSPropRMark", "sprmSPgnStart"]

PICT_PROPERTY_MODIFIER_VALUE = [0x6C02, 0x6C03, 0x6C04, 0x6C05, 0xCE08, 0xCE09, 0xCE0A, 0xCE0B]
PICT_PROPERTY_MODIFIER_NAME = ["sprmPicBrcTop80", "sprmPicBrcLeft80", "sprmPicBrcBottom80", "sprmPicBrcRight80", "sprmPicBrcTop", "sprmPicBrcLeft", "sprmPicBrcBottom", "sprmPicBrcRight"]



# PDF
# Extract Object 
REGULAR_EXPRESSION_OBJECT = ["obj\s{0,10}<<(.{0,100}?/Length\s([0-9]{1,8}).*?)>>[\s%]*?stream(.*?)endstream[^=]", "obj\s{0,10}<<(.{0,100}.*?)>>[\s%]*?stream(.*?)endstream[^=]"]
REGULAR_EXPRESSION_ESCAPE = "(\w*?)\s*?=\s*?(unescape\((.*?)\))"

# Condition JavaScript
g_JavaScript = ['/JS', '/JavaScript', '/FlateDecode', '/ASCII85Decode', '/ASCIIHexDecode', '/RunLengthDecode', '/LZWDecode']

# Decoding String & Decode Function
g_Decode = ['/FlateDecode', '/Fl', '/ASCII85Decode', '/ASCIIHexDecode', '/RunLengthDecode', '/LZWDecode']

# List Related Suspicious
g_SuspiciousList = ['Event', 'Action', 'JavaScript', 'Embedded']
g_SuspiciousEvents = ['/OpenAction', '/AA']
g_SuspiciousActions = ['/Launch', '/SubmitForm', '/ImportData']
g_SuspiciousJS = ['mailto', 
                  'Collab.collectEmailInfo', 
                  'util.printf', 
                  'getIcon', 
                  'getAnnots',                   
                  'spell.customDictionaryOpen', 
                  'media.newPlayer']

g_CVENo = ["CVE-2007-5020", 
           "CVE-2007-5659", 
           "CVE-2008-2992", 
           "CVE-2009-0927", 
           "CVE-2009-1492", 
           "CVE-2009-1493", 
           "CVE-2009-4324"]








