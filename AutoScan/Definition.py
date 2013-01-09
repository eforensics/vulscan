

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
g_Version = ["Office97", "Office2000", "Office2002", "Office2003", "Office2007"]


SIZE_OF_FIBFCLCB_97 = 0x2E8       # 744Bytes
RULE_FIBFCLCB_PATTERN_97 = '=186L'
RULE_FIBFCLCB_NAME_97 = ('fcStshfOrig lcbStshfOrig fcStshf lcbStshf fcPlcffndRef lcbPlcffndRef fcPlcffndTxt lcbPlcffndTxt fcPlcfandRef lcbPlcfandRef fcPlcfandTxt lcbPlcfandTxt fcPlcfSed   lcbPlcfSed fcPlcPad lcbPlcPad fcPlcfPhe   lcbPlcfPhe fcSttbfGlsy lcbSttbfGlsy fcPlcfGlsy lcbPlcfGlsy fcPlcfHdd lcbPlcfHdd fcPlcBteChpx lcbPlcBteChpx fcPlcfBtePapx   lcbPlcfBtePapx fcPlcfSea lcbPlcfSea fcSttbfFfn  lcbSttbfFfn fcPlcfFldMom lcbPlcfFldMom fcPlcfFldHdr lcbPlcfFldHdr fcPlcfFldFtn lcbPlcfFldFtn fcPlcfFldAtn lcbPlcfFldAtn fcPlcfFldMcr lcbPlcfFldMcr fcSttbfBkmk lcbSttbfBkmk fcPlcfBkf lcbPlcfBkf fcPlcfBkl lcbPlcfBkl fcCmds lcbCmds fcUnused1 lcbUnused1 fcSttbfMcr lcbSttbfMcr fcPrDrvr lcbPDrvr fcPrEnvPort lcbPrEnvPort fcPrEnvLand lcbPrEnvLand fcWss lcbWss fcDop lcbDop fcSttbfAssoc lcbSttbfAssoc fcClx lcbClx fcPlcfPgdFtn lcbPlcfPgdFtn fcAutosaveSource lcbAutosaveSource fcGrpXstAtnOwners lcbGrpXstAtnOwners fcSttbfAtnBkmk lcbSttbfAtnBkmk fcUnused3 lcbUnused3 fcUnused2 lcbUnused2 fcPlcSpaMon lcbPlcSpaMon fcPlcSpaHdr lcbPlcSpaHdr fcPlcAtnBkf lcbPlcAtnBkf fcPlcAtnBkl lcbPlcAtnBkl fcPms lcbPms fcFormFldSttbs  ldbFormFldSttbs fcPlcfendRef lcbPlcfendRef fcPlcfendTxt lcbPlcfendTxt fcPlcfFldEdn lcbPlcfFldEdn fcUnused4 lcbUnused4 fcDggInfo lcbDggInfo fcSttbfRMark lcbSttbfRMark fcSttbfCaption  lcbSttbfCaption fcSttbfAutoCaption  lcbSttbfAutoCaption fcPlcWkb lcbPlcWkb fcPlcfSpl lcbPlcfSpl fcPlcftxbxTxt lcbPlcftxbxTxt fcPlcfFldTxbx lcbPlcfFldTxbx fcPlcfHdrtxbxTxt lcbPlcfHdrtxbxTxt fcPlcffldHdrTxdx lcbPlcffldHdrTxdx fcStwUser lcbStwUser fcSttbTtmbd lcbSttbTtmbd fcCookieData lcbCookieData fcPgdMotherOldOld lcbPgdMotherOldOld fcBkdMotherOldOld lcbBkdMotherOldOld fcPgdFtnOldOld lcbPgdFtnOldOld fcBkdFtnOldOld lcbBkdFtnOldOld fcPgdEdnOldOld lcbPgdEdnOldOld fcBkdEdnOldOld lcbBkdEdnOldOld fcSttbfIntlFld lcbSttbfIntlFld fcRouteSlip lcbRouteSlip fcSttbSavedBy lcbSttbSavedBy fcSttbFnm lcbSttbFnm fcPlfLst lcbPlfLst fcPlfLfo lcbPlfLfo fcPlcTxbxBkd lcbPlcTxbxBkd fcPlcfTxbxHdrBkd lcbPlcfTxbxHdrBkd fcDocUndoWord9  lcbDocUndoWord9 fcRgbUse lcbRgbUse fcUsp lcbUsp fcUskf lcbUskf fcPlcupcRgbUse lcbPlcupcRgbUse fcPlcupcUsp lcbPlcupcUsp fcSttbGlsyStyle lcbSttbGlsyStyle fcPlgosl lcbPlgosl fcPlcocx lcbPlcocx fcPlcfBteLvc lcbPlcfBteLvc dwLowDateTime dwHighDateTime fcPlcfLvcPe10 lcbPlcfLvcPe10 fcPlcfAsumy lcbPlcfAsumy fcPlcfGram lcbPlcfGram fcSttbListNames lcbSttbListNames fcSttbfUssr lcbSttbfUssr')


SIZE_OF_FIBFCLCB_2000 = 0x78      # 120Bytes
RULE_FIBFCLCB_PATTERN_2000 = '=30L'
RULE_FIBFCLCB_NAME_2000 = ('fcPlcfTch lcbPlcfTch fcRmdThreading lcbRmdThreading fcMid lcbMid fcSttbRgtplc lcbSttbRgtplc fcMsoEnvelope lcbMsoEnvelope fcPlcfLad lcbPlcfLad fcRgDofr lcbRgDofr fcPlcosl lcbPlcosl fcPlcfCookieOld lcbPlcfCookieOld fcPgdMotherOld lcbPgdMotherOld fcBkdMotherOld lcbBkdMotherOld fcPgdFtnOld lcbPgdFtnOld fcBkdFtnOld lcbBkdFtnOld fcPgdEdnOld lcbPgdEdnOld fcBkdEdnOld lcbBkdEdnOld')


SIZE_OF_FIBFCLCB_2002 = 0xE0      # 224Bytes
RULE_FIBFCLCB_PATTERN_2002 = '=56L'
RULE_FIBFCLCB_NAME_2002 = ('fcUnused1 lcbUnused1 fcPlcfPgp lcbPlcfPgp fcPlcfuim lcbPlcfuim fcPlfguidUim lcbPlfguidUim fcAtrdExtra lcbAtrdExtra fcPlrsid lcbPlrsid fcSttbfBkmkFactoid lcbSttbfBkmkFactoid fcPlcfBkfFactoid lcbPlcfBkfFactoid fcPlcfcookie lcbPlcfcookie fcPlcfBklFactoid lcbPlcfBklFactoid fcFactoidData lcbFactoidData fcDocUndo lcbDocUndo fcSttbfBkmkFcc lcbSttbfBkmkFcc fcPlcfBkfFcc lcbPlcfBkfFcc fcPlcfBklFcc lcbPlcfBklFcc fcSttbfbkmkBPRepairs lcbSttbfbkmkBPRepairs fcPlcfbkfBPRepairs lcbPlcfbkfBPRepairs fcPlcfbklBPRepairs lcbPlcfbklBPRepairs fcPmsNew  lcbPmsNew fcODSO lcbODSO fcPlcfpmiOldXP lcbPlcfpmiOldXP fcPlcfpmiNewXP lcbPlcfpmiNewXP fcPlcfpmiMixedXP lcbPlcfpmiMixedXP fcUnused2 lcbUnused2 fcPlcffactoid lcbPlcffactoid fcPlcflvcOldXP lcbPlcflvcOldXP fcPlcflvcNewXP lcbPlcflvcNewXP fcPlcflvcMixedXP lcbPlcflvcMixedXP')


SIZE_OF_FIBFCLCB_2003 = 0xE0      # 224Bytes
RULE_FIBFCLCB_PATTERN_2003 = '=56L'
RULE_FIBFCLCB_NAME_2003 = ('fcHplxsdr lcbHplxsdr fcSttbfBkmkSdt lcbSttbfBkmkSdt fcPlcfBkfSdt lcbPlcfBkfSdt fcPlcfBklSdt lcbPlcfBklSdt fcCustomXForm lcbCustomXForm fcSttbfBkmkProt lcbSttbfBkmkProt fcPlcfBkfProt lcbPlcfBkfProt fcPlcfBklProt lcbPlcfBklProt fcSttbProtUser lcbSttbProtUser fcUnused lcbUnused fcPlcfpmiOld lcbPlcfpmiOld fcPlcfpmiOldInline,lcbPlcfpmiOldInline fcPlcfpmiNew lcbPlcfpmiNew fcPlcfpmiNewInline,lcbPlcfpmiNewInline fcPlcflvcOld lcbPlcflvcOld fcPlcflvcOldInline,lcbPlcflvcOldInline fcPlcflvcNew lcbPlcflvcNew fcPlcflcxNewInline,lcbPlcflcxNewInline fcPgdMother lcbPgdMother fcBkdMother lcbBkdMother fcAfdMother lcbAfdMother fcPgdFtn lcbPgdFtn fcBkdFtn lcbBkdFtn fcAfdFtn lcbAfdFtn fcPgdEdn lcbPgdEdn fcBkdEdn lcbBkdEdn fcAfdEdn lcbAfdEdn fcAfd lcbAfd')


SIZE_OF_FIBFCLCB_2007 = 0x98      # 152Bytes
RULE_FIBFCLCB_PATTERN_2007 = '=38L'
RULE_FIBFCLCB_NAME_2007 = ('fcPlcfmthd lcbPlcfmthd fcSttbfBkmkMoveFrom lcbSttbfBkmkMoveFrom fcPlcfBkfMoveFrom lcbPlcfBkfMoveFrom fcPlcfBklMoveFrom lcbPlcfBklMoveFrom fcSttbfBkmkMoveTo lcbSttbfBkmkMoveTo  fcPlcfBkfMoveTo lcbPlcfBkfMoveTo fcPlcfBklMoveTo lcbPlcfBklMoveTo fcUnused1 lcbUnused1 fcUnused2 lcbUnused2 fcUnused3 lcbUnused3 fcSttbfBkmkArto lcbSttbfBkmkArto fcPlcfBkfArto lcbPlcfBkfArto fcPlcfBklArto lcbPlcfBklArto fcArtoData lcbArtoData fcUnused4 lcbUnused4 fcUnused5 lcbUnused5 fcUnused6 lcbUnused6 fcOssTheme lcbOssTheme fcColorSchemeMapping lcbColorSchemeMapping')



















