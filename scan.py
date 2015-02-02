__author__ = 'KimSeungEon'

import os
import sys
import mmap
import struct
import traceback
from collections import namedtuple

def readUI8(buffer, offset):
    try:
        return struct.unpack("<B", buffer[offset])[0]
    except:
        return None

def readUI16(buffer, offset):
    try:
        return struct.unpack("<H", buffer[offset:offset + 2])[0]
    except:
        return None

def readUI24(buffer, offset):
    try:
        return struct.unpack("<L", buffer[offset:offset + 3] + "\x00")[0]
    except:
        return None

def readUI32(buffer, offset):
    try:
        return struct.unpack("<L", buffer[offset:offset + 4])[0]
    except:
        return None

def readUnsigned(buffer, offset, size):
    try:
        method_name_by_size = {
            1   :   readUI8,
            2   :   readUI16,
            3   :   readUI24,
            4   :   readUI32
        }
        return method_name_by_size[size](buffer, offset)
    except:
        return None

def readI8(buffer, offset):
    try:
        return struct.unpack("<b", buffer[offset])[0]
    except:
        return None

def readI16(buffer, offset):
    try:
        return struct.unpack("<h", buffer[offset:offset + 2])[0]
    except:
        return None

def readI24(buffer, offset):
    try:
        return struct.unpack("<i", "\x00" + buffer[offset:offset + 3])[0]
    except:
        return None

def readI32(buffer, offset):
    try:
        return struct.unpack("i", buffer[offset:offset + 4])[0]
    except:
        return None

def readSigned(buffer, offset, size):
    try:
        method_name_by_size = {
            1   :   readI8,
            2   :   readI16,
            3   :   readI24,
            4   :   readI32
        }
        return method_name_by_size[size](buffer, offset)
    except:
        return None

def Map(name, struct_buffer, struct_mem_name, struct_mem_size):
    struct_name = namedtuple(name, struct_mem_name)
    return struct_name._make( struct.unpack(struct_mem_size, struct_buffer))

#-----------------------------------------------------------------------------------------------------------------------
class PLC:
    def getN(self, Plcbuf, cbData):
        try:
            return (len(Plcbuf) - 4) / (4 + cbData)
        except:
            return -1

class PRL:
    """
    +-------------------------------+
    |              Prl              |
    +===============================+
    |           sprm(2Bytes)        |
    +-------------------------------+
    |        operand(variable)      |
    +-------------------------------+
    referred : page. 30 in [MS-DOC].pdf

    +-------------------------------+
    |             SPRM              |
    +===============================+
    |         ispmd(9Bits)          |
    +-------------------------------+
    |          fSpec(1Bit)          |
    +-------------------------------+
    |           sgc(3Bits)          |
    +-------------------------------+
    |           spra(3Bits)         |
    +-------------------------------+
    referred : page. 30 in [MS-DOC].pdf
    """
    def parseSprm(self, val):
        try:
            Sprm = {}
            Sprm["ispmd"] = val & 0x01FF
            Sprm["fSpec"] = (val / 512) & 0x0001
            Sprm["sgc"] = (val / 1024) & 0x0007
            Sprm["spra"] = (val / 8192)
            return True, Sprm
        except:
            return False, {}

    def parsePrl(self, buffer, start_offset):
        try:
            size = 0

            val = readUI16(buffer, start_offset)
            ret, Sprm = self.parseSprm( val )
            if ret == False:
                raise SystemError
            size += 2

            spra_list = [1, 1, 2, 4, 2, 2, "variable", 3]
            if spra_list[Sprm["spra"]] == "variable":
                # sprmTDefTable 0xD608  cb 2Bytes
                # sprmPChgTabs  0xC615  cb 1Bytes
                if val == 0xD608:
                    size += 2 + readUI16(buffer, start_offset + 2)
                else:
                    size += 1 + ord(buffer[start_offset + 2])
            else:
                size += spra_list[Sprm["spra"]]

            Prl = {}
            Prl["Sprm"] = Sprm
            return True, Prl, size
        except:
            return False, {}, -1

    def parseGrpPrl(self, buffer, start_offset, total_size):
        try:
            offset = start_offset

            Prls = []
            while offset < start_offset + total_size:
                if (start_offset + total_size) - offset < 3:    # Check Limit Size: 3
                    break
                ret, Prl, size = self.parsePrl(buffer, offset)
                if ret == False:
                    raise SystemError
                Prls.append(Prl)
                offset += size

            return True, Prls
        except:
            return False, []

class PnFkp:
    """
    refered : page. 433, 434 in [MS-DOC].pof

    +-------------------------------+
    | PnFkp( PnFkpChpx / PnFkpPapx )|
    +===============================+
    |           Pn(22Bits)          |-> referred WordDocument
    +-------------------------------+   (Pn * 512 = offset in WordDocument)
    |          unused(10Bits)       |
    +-------------------------------+
    """
    def parsePnFkp(self, buffer, start_offset, count):
        try:
            return True, [readUI32(buffer, start_offset + i * 4) & 0x3FFFFF for i in xrange(count)]
        except:
            return False, []
class PlcfBteChpx(PLC, PRL, PnFkp):
    """
    +-------------------------------+
    |          PlcfBteChpx          |
    +===============================+
    |          aFC(variable)        |-> referred WordDocument, Text
    +-------------------------------+
    |       aPnBteChpx(variable)    |-> PnFkpChpx
    +-------------------------------+
    referred : page. 197 in [MS-DOC].pdf

    +-------------------------------+
    |            ChpxFkp            |
    +===============================+
    |         rgfc(variable)        |-> referred WordDocument, Text
    +-------------------------------+
    |          rgb(variable)        |-> Chpx offset in WordDocument
    +-------------------------------+
    |           crun(1byte)         |
    +-------------------------------+
    referred : page. 237 in [MS-DOC].pdf

    +-------------------------------+
    |             Chpx              |
    +===============================+
    |           cb(1byte)           |
    +-------------------------------+
    |        grpprl(variable)       |-> Prls
    +-------------------------------+
    referred : page. 237 in [MS-DOC].pdf
    """
    def __init__(self, doc, table1, start_offset, total_size):
        self.doc = doc
        self.table = table1
        self.start_offset = start_offset
        self.total_size = total_size

        self.n = self.getN(self.table[self.start_offset:self.start_offset + self.total_size], 4)

    def parseChpxFkp(self, doc, Pn_list):
        try:
            chpxFkp = []
            for Pn in Pn_list:
                buf = doc[Pn * 512:(Pn + 1) * 512]              # ChpxFkp buffer in WordDocument
                crun = ord(buf[-1])
                offset = (crun + 1) * 4                         # Skip ChpxFkp.rgfc because BteChpx["Text"] is same text

                try:
                    Chpx_list = []                              # temp list-variable
                    for i in xrange(crun):
                        chpx_offset = ord(buf[offset+i]) * 2
                        cb = ord(buf[chpx_offset])              # Chpx.cb
                        ret, GrpPrl = self.parseGrpPrl(buf, chpx_offset + 1, cb)
                        if ret == False:
                            raise SystemError
                        Chpx_list.append(GrpPrl)

                    chpxFkp.append(Chpx_list)

                except:
                    raise SystemError

            return True, chpxFkp
        except:
            return False, []

    def parse(self):
        try:
            if self.n < 0:
                raise SystemError

            offset = self.start_offset

            tmp_list = []
            for i in xrange(self.n):
                tmp_list.append(self.doc[readUI32(self.table, offset + i * 4):readUI32(self.table, offset + (i+1) * 4):2])

            ret, Pn = self.parsePnFkp(self.table, offset + (self.n + 1) * 4, self.n)
            if ret == False:
                raise SystemError

            ret, chpxFkp = self.parseChpxFkp(self.doc, Pn)
            if ret == False:
                raise SystemError

            BteChpx = {}
            BteChpx["Text"] = tmp_list
            BteChpx["Pn"] = Pn
            BteChpx["Chpx"] = chpxFkp
            return True, BteChpx
        except:
            return False, {}
class PlcfBtePapx(PLC, PRL, PnFkp):
    """
    +-------------------------------+
    |          PlcfBtePapx          |
    +===============================+
    |          aFC(variable)        |-> referred WordDocument
    +-------------------------------+
    |       aPnBtePapx(variable)    |-> PnFkpPapx
    +-------------------------------+
    refered : page. 197 in [MS-DOC].pdf

    +-------------------------------+
    |           PapxFkp             |
    +===============================+
    |         rgfc(variable)        |-> referred WordDocument, Text
    +-------------------------------+
    |         rgbx(variable)        |-> BxPap Offset in WordDocument
    +-------------------------------+
    |          cpara(1Byte)         |
    +-------------------------------+
    refered : page. 413 in [MS-DOC].pdf

    +-------------------------------+
    |             BxPap             |
    +===============================+
    |          bOffset(1Byte)       |-> PapxInFkp in PapxFkp
    +-------------------------------+
    |        reserved(12Bytes)      |
    +-------------------------------+
    refered : page. 232 in [MS-DOC].pdf

    +-------------------------------+
    |           PapxInFkp           |
    +===============================+
    |           cb(1Byte)           |
    +-------------------------------+
    |      grpprlInPapx(variable)   |-> GrpPrlInIstd
    +-------------------------------+
    refered : page. 414 in [MS-DOC].pdf

    +-------------------------------+
    |           GrpPrlInIstd        |
    +===============================+
    |           istd(2Bytes)        |
    +-------------------------------+
    |       grpprl(variable)        |-> Prls
    +-------------------------------+
    refered : page. 370 in [MS-DOC].pdf
    """
    def __init__(self, doc, table1, start_offset, total_size):
        self.doc = doc
        self.table = table1
        self.start_offset = start_offset
        self.total_size = total_size

        self.n = self.getN(self.table[self.start_offset:self.start_offset + self.total_size], 4)

    def parsePapxFkp(self, doc, Pn_list):
        try:
            papxFkp = []
            for Pn in Pn_list:
                buf = doc[Pn * 512:(Pn + 1) * 512]              # PapxFkp buffer in WordDocument
                cpara = ord(buf[-1])
                offset = (cpara + 1) * 4                        # Skip PapxFkp.rgfc because BtePapx["Text"] is same text

                try:
                    Papx_list = []                              # temp list-variable
                    for i in xrange(cpara):
                        bOffset = ord(buf[offset + i * 13]) * 2
                        cb = ord(buf[bOffset])                  # PapxInFkp.cb
                        bOffset += 1
                        if cb != 0:
                            cb = 2 * cb - 1
                        else: # cb == 0
                            cb = ord(buf[bOffset])
                            bOffset += 1

                        istd = readUI16(buf, bOffset)           # GrpPrlAndIstd.istd
                        bOffset += 2

                        ret, GrpPrl = self.parseGrpPrl(buf, bOffset, cb)
                        if ret == False:
                            raise SystemError

                        Papx_list.append(GrpPrl)

                    papxFkp.append(Papx_list)

                except:
                    raise SystemError

            return True, papxFkp
        except:
            return False, []

    def parse(self):
        try:
            if self.n < 0:
                raise SystemError

            offset = self.start_offset

            tmp_list = []
            for i in xrange(self.n):
                tmp_list.append(self.doc[readUI32(self.table, offset + i * 4):readUI32(self.table, offset + (i+1) * 4):2])

            ret, Pn = self.parsePnFkp(self.table, offset + (self.n + 1) * 4, self.n)
            if ret == False:
                raise SystemError

            ret, papxFkp = self.parsePapxFkp(self.doc, Pn)
            if ret == False:
                raise SystemError

            BtePapx = {}
            BtePapx["Text"] = tmp_list
            BtePapx["Pn"] = Pn
            BtePapx["Papx"] = papxFkp
            return True, BtePapx
        except:
            return False, {}
class PlcfSed(PLC, PRL):
    """
    +-------------------------------+
    |            PlcfSed            |
    +===============================+
    |          aCP(variable)        |-> referred WordDocument, Text
    +-------------------------------+
    |          aSed(variable)       |-> Sed
    +-------------------------------+
    referred : page. 209 in [MS-DOC].pdf

    +-------------------------------+
    |              Sed              |
    +===============================+
    |           fn(2Bytes)          |-> Ignore
    +-------------------------------+
    |         fcSepx(4Bytes)        |-> Sepx, referred WordDocument
    +-------------------------------+
    |         fnMpr(2Bytes)         |-> Ignore
    +-------------------------------+
    |         fcMpr(4Bytes)         |-> Ignore
    +-------------------------------+
    referred : page. 456 ~ 457 in [MS-DOC].pdf

    +-------------------------------+
    |             Sepx              |
    +===============================+
    |           cb(2Bytes)          |
    +-------------------------------+
    |         grpprl(variable)      |-> Prl
    +-------------------------------+
    referred : page. 459 ~ 460 in [MS-DOC].pdf
    """
    def __init__(self, doc, table1, start_offset, total_size):
        self.doc = doc
        self.table = table1
        self.start_offset = start_offset
        self.total_size = total_size

        self.n = self.getN(self.table[self.start_offset:self.start_offset + self.total_size], 12)

    def parse(self):
        try:
            if self.n < 0:
                raise SystemError

            offset = self.start_offset

            tmp_list = []
            for i in xrange(self.n):
                tmp_list.append(self.doc[readUI32(self.table, offset + i * 4):readUI32(self.table, offset + (i+1) * 4):2])

            tmp_sepx = []
            offset += (self.n + 1) * 4
            for i in xrange(self.n):
                Sepx_offset = readI32(self.table, offset + i * 4 + 2)     # Sepx Offset in WordDocument
                cb = readI16(self.doc, Sepx_offset)
                ret, GrpPrl = self.parseGrpPrl(self.doc, Sepx_offset+2, cb)
                if ret == False:
                    raise SystemError
                tmp_sepx.append(GrpPrl)

            plcfSed = {}
            plcfSed["Text"] = tmp_list
            plcfSed["Sepx"] = tmp_sepx
            return True, plcfSed
        except:
            return False, {}

class Clx(PLC, PRL):
    """"
    +-------------------------------+
    |              Clx              |
    +===============================+
    |         RgPrc(variable)       |-> Prc
    +-------------------------------+
    |         Pcdt(variable)        |-> Pcdt
    +-------------------------------+
    referred : page. 242 in [MS-DOC].pdf

    +-------------------------------+
    |              Prc              |
    +===============================+
    |           clxt(1Byte)         | MUST 0x01
    +-------------------------------+
    |         data(variable)        |-> PrcData
    +-------------------------------+
    referred : page. 434 ~ 435 in [MS-DOC].pdf

    +-------------------------------+
    |            PrcData            |
    +===============================+
    |       cbGrpprl(2Bytes)        | <= 0x3FA2
    +-------------------------------+
    |       GrpPrl(variable)        |-> Prls
    +-------------------------------+
    referred : page. 435 in [MS-DOC].pdf

    +-------------------------------+
    |             Pcdt              |
    +===============================+
    |           clxt(1Byte)         | MUST 0x02
    +-------------------------------+
    |           lcb(4Bytes)         |
    +-------------------------------+
    |         PlcPcd(variable)      |-> PlcPcd
    +-------------------------------+
    referred : page. 415 in [MS-DOC].pdf

    +-------------------------------+
    |             PlcPcd            |
    +===============================+
    |         aCP(variable)         |
    +-------------------------------+
    |        aPcd(variable)         |-> Pcd
    +-------------------------------+
    referred : page. 214 ~ 215 in [MS-DOC].pdf

    +-------------------------------+
    |              Pcd              |
    +===============================+
    |       fNoParaLast(1Bit)       |
    +-------------------------------+
    |           fR1(1Bit)           | Ignore
    +-------------------------------+
    |         fDirty(1Bit)          | MUST 0
    +-------------------------------+
    |           fR2(13Bits)         | Ignore
    +-------------------------------+
    |           fc(4Bytes)          |-> FcCompressed
    +-------------------------------+
    |           prm(2Bytes)         |-> Prm
    +-------------------------------+
    referred : page. 415 in [MS-DOC].pdf

    +-------------------------------+
    |        FcCompressed           |
    +===============================+
    |           fc(30Bits)          | if fCompressed == 0: fc else fc / 2
    +-------------------------------+
    |         fCompressed(1Bit)     |
    +-------------------------------+
    |            r1(1Bit)           | MUST 0, Ignore
    +-------------------------------+
    referred : page. 262 in [MS-DOC].pdf

    +-------------------------------+
    |              Prm              |
    +===============================+
    |         fComplex(1Bit)        |
    +-------------------------------+
    |           Data(15Bits)        |-> if fComplex == 0: Prm0 else Prm1
    +-------------------------------+
    referred : page. 436 in [MS-DOC].pdf

    +-------------------------------+
    |               Prm0            |
    +===============================+
    |           isprm(7Bits)        |-> referred Sprm
    +-------------------------------+
    |           val(8Bits)          |
    +-------------------------------+
    referred : page. 436 ~ 437 in [MS-DOC].pdf

    +-------------------------------+
    |               Prm1            |
    +===============================+
    |         igrpprl(15Bits)       |
    +-------------------------------+
    referred : page. 438 ~ 439 in [MS-DOC].pdf
    """
    def __init__(self, doc, table1, start_offset, total_size):
        self.doc = doc
        self.table = table1
        self.start_offset = start_offset
        self.total_size = total_size

    def _getclxt(self, table, offset):
        try:
            clxt = ord(table[offset])
            if not clxt in [0x01, 0x02]:
                raise SystemError
            return True, clxt
        except:
            return False, {}

    def _parsePlcPcd(self, table, start_offset, total_size):
        try:
            n = self.getN(total_size, 8)
            offset = start_offset + (n + 1) * 4     # skip aCP

            PlcPcd = {}
            tmp = readUI32(table, offset + 2)
            PlcPcd["fCompressed"] = (tmp >> 30) & 0x1
            if PlcPcd["fCompressed"] == 0:
                PlcPcd["fc"] = tmp & 0x3FFF
            else:
                PlcPcd["fc"] = (tmp & 0x3FFF) / 2

            tmp = readUI16(table, offset + 6)
            if tmp & 0x1 == 0:  # Prm0
                PlcPcd["isprm"] = (tmp >> 1) & 0x7F
            else: # Prm1
                PlcPcd["igrpprl"] = tmp >> 1

            return True, PlcPcd
        except:
            return False, {}

    def parse(self):
        try:
            offset = self.start_offset

            ret, clxt = self._getclxt(self.table, offset)
            if ret == False:
                raise SystemError
            offset += 1

            clx = {}
            if clxt == 0x01:
                cbGrpPrl = readUI16(self.table, offset)
                if not (cbGrpPrl <= 0x3FA2):
                    raise SystemError
                offset += 2
                ret, GrpPrl = self.parseGrpPrl(self.table, offset, cbGrpPrl)
                if ret == False:
                    raise SystemError
                offset += cbGrpPrl
                clx["RgPrc"] = GrpPrl

            if clx != {}:       # Parsed clx["RgPrc"]
                ret, clxt = self._getclxt(self.table, offset)
                if ret == False:
                    raise SystemError
                offset += 1

            if clxt == 0x02:
                ret, PlcPcd = self._parsePlcPcd(self.table, offset + 4, readUI32(self.table, offset))
                if ret == False:
                    raise SystemError
                clx["PlcPcd"] = PlcPcd

            return True, clx
        except:
            return False, {}



class FibRgFcLcb97:
    def __init__(self, doc, table, start_offset, size):
        self.doc = doc
        self.table = table
        self.total_size = size

        self.FibRgFcLcb97_start = start_offset
        self.FibRgFcLcb97_Size = 744

    def parse(self):
        try:
            str97_List = ["fcStshfOrig", "Stshf", "PlcffndRef", "PlcffndTxt", "PlcfandRef", "PlcfandTxt", "PlcfSed", "PlcPad", "PlcfPhe", "SttbfGlsy", "PlcfGlsy", "PlcfHdd", "PlcfBteChpx", "PlcfBtePapx", "PlcfSea", "SttbfFfn", "PlcfFldMom", "PlcfFldHdr", "PlcfFldFtn", "PlcfFldAtn", "PlcfFldMcr", "SttbfBkmk", "PlcfBkf", "PlcfBkl", "Cmds", "Unused1", "SttbfMcr", "PrDrvr", "PrEnvPort", "PrEnvLand", "Wss", "Dop", "SttbfAssoc", "Clx", "PlcfPgdFtn", "AutosaveSource", "GrpXstAtnOwners", "SttbfAtnBkmk", "Unused2", "Unused3", "PlcSpaMom", "PlcSpaHdr", "PlcfAtnBkf", "PlcfAtnBkl", "Pms", "FormFldSttbs", "PlcfendRef", "PlcfendTxt", "PlcfFldEdn", "Unused4", "DggInfo", "SttbfRMark", "SttbfCaption", "SttbfAutoCaption", "PlcfWkb", "PlcfSpl", "PlcftxbxTxt", "PlcfFldTxbx", "PlcfHdrtxbxTxt", "PlcffldHdrTxbx", "StwUser", "SttbTtmbd", "CookieData", "PgdMotherOldOld", "BkdMotherOldOld", "PgdFtnOldOld", "BkdFtnOldOld", "PgdEdnOldOld", "BkdEdnOldOld", "SttbfIntlFld", "RouteSlip", "SttbSavedBy", "SttbFnm", "PlfLst", "PlfLfo", "PlcfTxbxBkd", "PlcfTxbxHdrBkd", "DocUndoWord9", "RgbUse", "Usp", "Uskf", "PlcupcRgbUse", "PlcupcUsp", "SttbGlsyStyle", "Plgosl", "Plcocx", "PlcfBteLvc", "DateTime", "PlcfLvcPre10", "PlcfAsumy", "PlcfGram", "SttbListNames", "SttbfUssr"]

            struct97 = {}
            offset = self.FibRgFcLcb97_start
            for funcName in str97_List:
                try:
                    pos = readUI32(self.doc, offset)
                    size = readUI32(self.doc, offset+4)
                    inst_class = getattr(sys.modules[__name__], funcName)(self.doc, self.table, pos, size)
                    ret, str97 = inst_class.parse()
                    if ret == False:
                        raise SystemError

                    struct97[funcName] = str97

                except SystemError:
                    raise SystemError

                except AttributeError:
                    pass

                except:
                    print traceback.format_exc()

                finally:
                    offset += 8

            return True, struct97
        except:
            return False, {}

class FibRgFcLcb2000(FibRgFcLcb97):
    def __init__(self, doc, table, start_offset, size):
        FibRgFcLcb97.__init__(self, doc, table, start_offset, size)
        struct97 = FibRgFcLcb97.parse(self)

        self.FibRgFcLcb2000_Start = self.FibRgFcLcb97_start + self.FibRgFcLcb97_Size
        self.FibRgFcLcb2000_Size = 120

    def parse(self):
        try:
            str2000_List = ["PlcfTch", "RmdThreading", "Mid", "SttbRgtplc", "MsoEnvelope", "PlcfLad", "RgDofr", "Plcosl", "PlcfCookieOld ", "PgdMotherOld", "BkdMotherOld", "PgdFtnOld", "BkdFtnOld", "PgdEdnOld", "BkdEdnOld"]

            struct2000 = {}
            offset = self.FibRgFcLcb2000_Start
            for funcName in str2000_List:
                try:
                    pos = readUI32(self.doc, offset)
                    size = readUI32(self.doc, offset+4)
                    inst_class = getattr(sys.modules[__name__], funcName)(table, pos, size)
                    ret, str2000 = inst_class.parse()
                    if ret == False:
                        raise SystemError

                    struct2000[funcName] = str2000

                except SystemError:
                    raise SystemError

                except:
                    pass

                finally:
                    offset += 8

            return True, struct2000
        except:
            return False, {}

class FibRgFcLcb2002(FibRgFcLcb2000):
    def __init__(self, doc, table, start_offset, size):
        FibRgFcLcb2000.__init__(self, doc, table, start_offset, size)
        struct2000 = FibRgFcLcb2000.parse(self)

        self.FibRgFcLcb2002_Start = self.FibRgFcLcb2000_Start + self.FibRgFcLcb2000_Size
        self.FibRgFcLcb2002_Size = 0

    def parse(self):
        try:
            str2002_List = ["Unused1", "PlcfPgp", "Plcfuim", "PlfguidUim", "AtrdExtra", "Plrsid", "SttbfBkmkFactoid", "PlcfBkfFactoid", "Plcfcookie", "PlcfBklFactoid", "FactoidData", "DocUndo", "SttbfBkmkFcc", "PlcfBkfFcc", "PlcfBklFcc", "SttbfbkmkBPRepairs", "PlcfbkfBPRepairs", "PlcfbklBPRepairs", "PmsNew", "ODSO", "PlcfpmiOldXP", "PlcfpmiNewXP", "PlcfpmiMixedXP", "Unused2", "Plcffactoid", "PlcflvcOldXP", "PlcflvcNewXP", "PlcflvcMixedXP"]

            struct2002 = {}
            offset = self.FibRgFcLcb2002_Start
            for funcName in str2002_List:
                try:
                    pos = readUI32(self.doc, offset)
                    size = readUI32(self.doc, offset+4)
                    inst_class = getattr(sys.modules[__name__], funcName)(table, pos, size)
                    ret, str2002 = inst_class.parse()
                    if ret == False:
                        raise SystemError

                    struct2002[funcName] = str2002

                except SystemError:
                    raise SystemError

                except:
                    pass

                finally:
                    offset += 8

            return True, struct2002
        except:
            return False, {}

class FibRgFcLcb2003(FibRgFcLcb2002):
    def __init__(self, doc, table, start_offset, size):
        FibRgFcLcb2002.__init__(self, doc, table, start_offset, size)
        struct2002 = FibRgFcLcb2002.parse(self)

        self.FibRgFcLcb2003_Start = self.FibRgFcLcb2002_Start + self.FibRgFcLcb2002_Size
        self.FibRgFcLcb2003_Size = 0

    def parse(self):
        try:
            str2003_List = ["Hplxsdr", "SttbfBkmkSdt", "PlcfBkfSdt", "PlcfBklSdt", "CustomXForm", "SttbfBkmkProt", "PlcfBkfProt", "PlcfBklProt", "SttbProtUser", "Unused", "PlcfpmiOld", "PlcfpmiOldInline", "PlcfpmiNew", "PlcfpmiNewInline", "PlcflvcOld", "PlcflvcOldInline", "PlcflvcNew", "PlcflvcNewInline", "PgdMother", "BkdMother", "AfdMother", "PgdFtn", "BkdFtn", "AfdFtn", "PgdEdn", "BkdEdn", "AfdEdn", "Afd"]

            struct2003 = {}
            offset = self.FibRgFcLcb2003_Start
            for funcName in str2003_List:
                try:
                    pos = readUI32(self.doc, offset)
                    size = readUI32(self.doc, offset+4)
                    inst_class = getattr(sys.modules[__name__], funcName)(table, pos, size)
                    ret, str2003 = inst_class.parse()
                    if ret == False:
                        raise SystemError

                    struct2003[funcName] = str2003

                except SystemError:
                    raise SystemError

                except:
                    pass

                finally:
                    offset += 8

            return True, struct2003
        except:
            return False, {}

class FibRgFcLcb2007(FibRgFcLcb2003):
    def __init__(self, doc, table, start_offset, size):
        FibRgFcLcb2003.__init__(self, doc, table, start_offset, size)
        struct2003 = FibRgFcLcb2003.parse(self)

        self.FibRgFcLcb2007_Start = self.FibRgFcLcb2003_Start + self.FibRgFcLcb2003_Size
        self.FibRgFcLcb2007_Size = 0

    def parse(self):
        try:
            str2007_List = ["Plcfmthd", "SttbfBkmkMoveFrom", "PlcfBkfMoveFrom", "PlcfBklMoveFrom", "SttbfBkmkMoveTo", "PlcfBkfMoveTo", "PlcfBklMoveTo", "Unused1", "Unused2", "Unused3", "SttbfBkmkArto", "PlcfBkfArto", "PlcfBklArto", "ArtoData", "Unused4", "Unused5", "Unused6", "OssTheme", "ColorSchemeMapping"]
            struct2007 = {}
            offset = self.FibRgFcLcb2007_Start
            for funcName in str2007_List:
                try:
                    pos = readUI32(self.doc, offset)
                    size = readUI32(self.doc, offset+4)
                    inst_class = getattr(sys.modules[__name__], funcName)(table, pos, size)
                    ret, str2007 = inst_class.parse()
                    if ret == False:
                        raise SystemError

                    struct2007[funcName] = str2007

                except SystemError:
                    raise SystemError

                except:
                    pass

                finally:
                    offset += 8

            return True, struct2007
        except:
            return False, {}

class FIB:
    def __init__(self, word, table):
        self.word  = word
        self.table = table

    def __uninit__(self):
        if self.word  != "":  self.word  = ""
        if self.table != "":  self.table = ""

    def _getnFib(self):
        try:
            base_nFib = readUI16(self.word, 2)
            cswNew_offset = 154 + readUI16(self.word, 152) * 8
            cswNew = readUI16(self.word, cswNew_offset)
            nFibNew = readUI16(self.word, cswNew_offset + 2)
            if cswNew:
                nFib = nFibNew
            else:
                nFib = base_nFib

            return True, nFib
        except:
            return False, -2

    def _FibBase(self):
        try:
            mem_name = ("wIdent nFib unused lid pnNext Flag1 nFibBack lKey Envr Flag2 reserved3 reserved4 reserved5 reserved6")
            mem_size = "<7H1L2B2H2L"
            if struct.calcsize(mem_size) != 32:
                raise SystemError

            self.FibBaseInfo = (0, struct.calcsize(mem_size))
            return True, Map("FibBase", self.word[self.FibBaseInfo[0]:self.FibBaseInfo[1]], mem_name, mem_size)
        except:
            return False, {}

    def _FibRgW(self):
        try:
            start_offset = self.FibBaseInfo[0] + self.FibBaseInfo[1]

            mem_name = ("reserved1 reserved2 reserved3 reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 lidFE")
            mem_size = "<14H"

            csw = readUI16(self.word, start_offset) * 2
            if csw != struct.calcsize(mem_size) != 28:
                raise SystemError

            self.FibRgWInfo = (start_offset + 2, struct.calcsize(mem_size))
            return True, Map("FibRgW", self.word[self.FibRgWInfo[0]:self.FibRgWInfo[0] + self.FibRgWInfo[1]], mem_name, mem_size)
        except:
            return False, {}

    def _FibRgLw(self):
        try:
            start_offset = self.FibRgWInfo[0] + self.FibRgWInfo[1]

            mem_name = ("cbMac reserved1 reserved2 ccpText ccpFtn ccpHdd reserved3 ccpAtn ccpEdn ccpTxbx ccpHdrTxbx reserved4 reserved5 reserved6 reserved7 reserved8 reserved9 reserved10 reserved11 reserved12 reserved13 reserved14")
            mem_size = "<3L3i1L4i11L"

            cslw = readUI16(self.word, start_offset) * 4
            if cslw != struct.calcsize(mem_size) != 88:
                raise SystemError

            self.FibRgLwInfo = (start_offset + 2, struct.calcsize(mem_size))
            FibRgLw = Map("FibRgLw", self.word[self.FibRgLwInfo[0]:self.FibRgLwInfo[0] + self.FibRgLwInfo[1]], mem_name, mem_size)

            if FibRgLw.ccpText < 0 or FibRgLw.ccpFtn < 0 or FibRgLw.ccpHdd < 0 or FibRgLw.ccpAtn < 0 or \
                FibRgLw.ccpEdn < 0 or FibRgLw.ccpTxbx < 0 or FibRgLw.ccpHdrTxbx < 0:
                raise SystemError

            return True, FibRgLw
        except:
            return False, {}

    def _FibRgFcLcb(self):
        try:
            start_offset = self.FibRgLwInfo[0] + self.FibRgLwInfo[1]
            cbRgFcLcb = readUI16(self.word, start_offset) * 8
            self.FibRgFcLcbInfo = (start_offset + 2, cbRgFcLcb)

            Blob_List = {
                0x00C1  :   "FibRgFcLcb97",
                0x00D9  :   "FibRgFcLcb2000",
                0x0101  :   "FibRgFcLcb2002",
                0x010C  :   "FibRgFcLcb2003",
                0x0112  :   "FibRgFcLcb2007",
            }

            inst_class = getattr(sys.modules[__name__], Blob_List[self.nFib])(self.word, self.table, self.FibRgFcLcbInfo[0], self.FibRgFcLcbInfo[1])
            ret, FibRgFcLcb = inst_class.parse()
            if ret == False:
                raise SystemError

            return True, FibRgFcLcb
        except:
            return False, {}

    def _FibRgCswNew(self):
        def FibRgCswNew2000(buf, start_offset, size):
            try:
                mem_name = ("cQuickSavesNew")
                mem_size = "<1H"
                if struct.calcsize(mem_size) != size != 2:
                    raise SystemError

                return True, Map("FibRgCswNew2000", buf[start_offset:start_offset + size], mem_name, mem_size)
            except:
                return False, {}

        def FibRgCswNew2007(buf, start_offset, size):
            try:
                mem_name = ("cQuickSavesNew lidThemeOther lidThemeFE lidThemeCS")
                mem_size = "<4H"
                if struct.calcsize(mem_size) != size != 8:
                    raise SystemError

                return True, Map("FibRgCswNew2007", buf[start_offset:start_offset + size], mem_name, mem_size)
            except:
                return False, {}

        try:
            start_offset = self.FibRgFcLcbInfo[0] + self.FibRgFcLcbInfo[1]
            cswNew = readUI16(self.word, start_offset)
            cswNew_List = {
                0x00C1  :   0x0000,
                0x00D9  :   0x0002,
                0x0101  :   0x0002,
                0x010C  :   0x0002,
                0x0112  :   0x0005
            }
            if cswNew != cswNew_List[self.nFib]:
                raise SystemError

            nFibNew = readUI16(self.word, start_offset + 2)
            CswNewData = {
                0x00D9  :   (FibRgCswNew2000, 2),
                0x0101  :   (FibRgCswNew2000, 2),
                0x010C  :   (FibRgCswNew2000, 2),
                0x0112  :   (FibRgCswNew2007, 8)
            }

            ret, FibRgCswNew = CswNewData[nFibNew][0](self.word, start_offset, CswNewData[nFibNew][1])
            if ret == False:
                raise SystemError

            return True, FibRgCswNew
        except:
            print traceback.format_exc()
            return False, {}

    def parseFIB(self):
        try:
            ret, self.nFib = self._getnFib()
            if ret == False:
                raise SystemError

            ret, FibBase = self._FibBase()
            if ret == False:
                raise SystemError

            ret, FibRgW = self._FibRgW()
            if ret == False:
                raise SystemError

            ret, FibRgLw = self._FibRgLw()
            if ret == False:
                raise SystemError

            ret, FibRgFcLcb = self._FibRgFcLcb()
            if ret == False:
                raise SystemError

            ret, FibRgCswNew = self._FibRgCswNew()
            if ret == False:
                raise SystemError

            FIB = {}
            FIB["base"] = FibBase
            FIB["FibRgW97"] = FibRgW
            FIB["FibRgLw97"] = FibRgLw
            FIB["FibRgFcLcb"] = FibRgFcLcb
            FIB["FibRgCswNew"] = FibRgCswNew
            return True, FIB
        except:
            return False, {}

class POLICY:
    def __init__(self):
        Policy = self._loadpolicy()

    def _loadpolicy(self):
        pass

class WORDOCUMENT(FIB):
    def __init__(self, word_buf, table_buf):
        FIB.__init__(self, word_buf, table_buf)

    def __uninit__(self):
        FIB.__uninit__()

    def scan(self):
        try:
            ret, FIB = self.parseFIB()
            if ret == False:
                return False, -2

            # TODO :: Check Pattern

            return True, 0
        except:
            return False, -1

if __name__ == "__main__":
    scan_path = r"C:\Users\KimSeungEon\Desktop\CVE\CVE-2008-2244\Exploit.CVE-2008-2244.Gen"
    wordDocument = "WordDocument.stream"
    table = "1Table.stream"

    fp = open(scan_path + os.sep + wordDocument)
    word = mmap.mmap(fp.fileno(), 0, access=mmap.ACCESS_READ)
    fp.close()

    fp = open(scan_path + os.sep + table)
    table = mmap.mmap(fp.fileno(), 0, access=mmap.ACCESS_READ)
    fp.close()

    output_list = {
        1       :       "CVE-2006",
        0       :       "Normal",
        -1      :       "scan()",
        -2      :       "parseFIB()"
    }

    print "[+] Scan ( %s )" % scan_path.split(os.sep)[-1]

    inst_class = WORDOCUMENT(word, table)
    ret, msg_id = inst_class.scan()
    if ret == False:
        if msg_id < 0:
            print "[-] Error : {}".format(output_list[msg_id])
        else:
            print "[+] Found : {}".format(output_list[msg_id])
    else:
        print "[+] Not Found"

    print "[+] END"

    inst_class.__uninit__()
    del inst_class
    sys.exit()