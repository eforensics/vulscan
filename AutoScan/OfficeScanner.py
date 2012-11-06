# -*- coding:utf-8 -*-

# import Public Module
import traceback, sys, os, binascii

# import Private Module
import Common



class Office():
    @classmethod
    def Scan(cls, File):
        try :
            Office = OperateOffice()
            
            # Extract Sector Data
            if not Office.ExtractSector( File ) :
                return False
            
            # Parse WordDocument ( FIB : File Information Block )
            MapOffice = MappedOffice()
            pWordDocument = File["WordDocument"]
            if pWordDocument == "" :
                print "\t\t[ERROR] WordDocument is NOT Saved"
                return False
            
            MapFIB = MapOffice.MappedFIB( pWordDocument )
            
            print MapFIB
            
            
            
            
            
            return True
        
        except :
            print traceback.format_exc()
            return False
        



class OperateOffice():
    def ExtractSector(self, File):
        try : 
            pBuf = File["pBuf"]
            DirList = File["DirList"]
            ExtType = []
            
            # Extract Sector By SAT
            SATable = File["SATable"]
            RefSAT = File["RefSAT"]
            for RefEntry in RefSAT :
                for DirEntry in DirList :
                    if RefEntry == DirEntry["EntryName"] :
                        Sector = self.ExtractSectorbySAT(pBuf, SATable, DirEntry["SecID"], DirEntry["szData"])
                        fname = "%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"])
                        Common.FileControl.WriteFile(fname, Sector)
                        
                        if RefEntry == "WordDocument" : 
                            File["WordDocument"] = Sector
                        
                        if RefEntry == "RootEntry" or RefEntry == "ootEntry" :
                            File["RootEntry"] = Sector                        
                        
                        if RefEntry == "FileHeader" :
                            File["FileHeader"] = Sector
                        
                        if RefEntry == "DocInfo" or (RefEntry.find("BinaryData") != -1) or (RefEntry.find("VersionLog") != -1) or (RefEntry.find("Section") != -1) :
                            ExtType.append(fname)
                        
                        break
                
            # Extract Sector By SSAT
            pBuf = File["RootEntry"]
            if pBuf == "" :
                print "\t\t[ERROR] RootEntry is NOT Saved"
                return False
            
            SSATable = File["SSATable"]
            RefSSAT = File["RefSSAT"]
            for RefEntry in RefSSAT :
                for DirEntry in DirList :
                    if RefEntry == "" :
                        continue                    
                        
                    if RefEntry == DirEntry["EntryName"] :
                        Sector = self.ExtractSectorbySSAT(pBuf, SSATable, DirEntry["SecID"], DirEntry["szData"])
                        fname = "%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"])
                        Common.FileControl.WriteFile(fname, Sector)
                        
                        if RefEntry == "WordDocument" : 
                            File["WordDocument"] = Sector
                            
                        if RefEntry == "FileHeader" :
                            File["FileHeader"] = Sector
                            
                        if RefEntry == "DocInfo" or (RefEntry.find("BinaryData") != -1) or (RefEntry.find("VersionLog") != -1) or (RefEntry.find("Section") != -1) :
                            ExtType.append(fname)
                            
                        break
            
            File["ExtType"] = ExtType
            
        except :
            print traceback.format_exc()
            return False
            
        return True
        
        
    def ExtractSectorbySAT(self, pBuf, Table, SecID, Size):
        try :
            Sector = ""
            tmpSector = ""
            while True :
                if SecID == 0xffffffff or SecID == 0xfffffffe or SecID == 0xfffffffd or SecID == 0xfffffffc :
                    break
            
                tmpSector += Common.BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x200)
                SecID = Table[SecID]
            
            Sector = Common.BufferControl.Read(tmpSector, 0, Size)
        
        except :
            print traceback.format_exc()
            return ""
            
        return Sector


    def ExtractSectorbySSAT(self, pBuf, Table, SecID, Size):
        try :
            Sector = ""
            tmpSector = ""
            while True :
                if SecID == 0xffffffff or SecID == 0xfffffffe or SecID == 0xfffffffd or SecID == 0xfffffffc :
                    break
            
                tmpSector += Common.BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x40)
                SecID = Table[SecID]
            
            Sector = Common.BufferControl.Read(tmpSector, 0, Size)
            
        except :
            print traceback.format_exc()
            return ""
        
        return Sector
        

class MappedOffice():
    def MappedFIB(self, pWordDocument):
        try :
            MapFIB = {}
            OutMapFIB = {}
            Offset = 0
            
            print "\t   [-] Step1. Mapped FibBase"
            # Step1. Mapped FibBase
            #    - Offset   : 0x0000
            #    - Size     : 0x0020 ( 32bytes )
            #    - By Per   : variable
            FibBase = self.MappedFibBase(pWordDocument)
            if FibBase == {} :
                return OutMapFIB
            
            if FibBase["wIdent"] != 0xA5EC :
                print "\t\t[ERROR] MappedFibBase( %s )" % FibBase["wIdent"]
                return OutMapFIB
            
            Offset += szFibBase
            
            print "\t   [-] Step2-1. Read csw"
            # Step2-1. Read csw
            #    - Offset    : 0x0020
            #    - Size      : 0x0002 ( 2bytes )
            csw = Common.BufferControl.ReadWord(pWordDocument, Offset)
            if csw != 0x000E :
                print "\t\t[ERROR] FibRgW97's csw ( 0x%04X : 0x%04X )" % ( Offset, csw )
                return OutMapFIB
            
            Offset += 2
            
            print "\t   [-] Step2-2. Mapped FibRgW97"
            # Step2-2. Mapped FibRgW97
            #    - Offset    : 0x0022
            #    - Size      : 0x001C ( 30bytes )
            #    - By Per    : 2bytes 
            szFibRgW97 = csw * 2
            FibRgW97 = self.MappedFibRgW97(pWordDocument, Offset, szFibRgW97)
            if FibRgW97 == {} :
                return OutMapFIB
            
            Offset += szFibRgW97
            
            print "\t   [-] Step3-1. Read cslw"
            # Step3-1. Read cslw
            #    - Offset    : 0x003E
            #    - Size      : 0x0002 ( 2bytes )
            cslw = Common.BufferControl.ReadWord(pWordDocument, Offset)
            if cslw != 0x0016 :
                print "\t\t[ERROR] FibRgLw97's cslw ( 0x%04X : 0x%04X )" % ( Offset, cslw )
                return OutMapFIB
            
            Offset += 2
            
            print "\t   [-] Step3-2. Mapped FibRgLw97"
            # Step3-2. Mapped FibRgLw97
            #    - Offset    : 0x0040
            #    - Size      : 0x0058 ( 88bytes )
            #    - By Per    : 4bytes
            szFibRgLw97 = cslw * 4
            FibRgLw97 = self.MappedFibRgLw97(pWordDocument, Offset, szFibRgLw97)
            if FibRgLw97 == {} :
                return OutMapFIB
            
            Offset += szFibRgLw97
            
            print "\t   [-] Step4-1. Read cdRgFcLcb"
            # Step4-1. Read cbRgFcLcb
            #    - Offset    : 0x0098
            #    - Size      : 0x0002 ( 2bytes )
            cbRgFcLcb = Common.BufferControl.ReadWord(pWordDocument, Offset)
            if cbRgFcLcb <= 0 :
                print "\t\t[ERROR] FibRgFcLcbBlob's cbRgFcLcb ( 0x%04X : 0x%04X )" % ( Offset, cbRgFcLcb )
                return OutMapFIB
            
            Offset += 2
            
            print "\t   [-] Step4-2. Mapped FibRgFcLcbBlob"
            # Step4-2. Mapped FibRgFcLcbBlob
            #    - Offset    : 0x009A
            #    - Size      : Variable
            #        FibRgFcLcb97    : 0x02E8 ( 744bytes )  
            #        FibRgFcLcb2000  : 0x0360 ( 864bytes )
            #        FibRgFcLcb2002  : 0x0440 ( 1088bytes )
            #        FibRgFcLcb2003  : 0x0520 ( 1312bytes )  
            #        FibRgFcLcb2007  : 0x05B8 ( 1464bytes )
            #    - By Per    : 2bytes 
            szFibRgFcLcbBlob = cbRgFcLcb * 8
            FibRgFcLcb = self.MappedFibRgFcLcbBlob(pWordDocument, Offset, FibBase["nFib"])
            if FibRgFcLcb == {} :
                return OutMapFIB
            
            MapFIB = FibRgFcLcb
                    
            Offset += szFibRgFcLcbBlob
            
            print "\t   [-] Step 5-1. Read cswNew"
            # Step5-1. Read cswNew
            #    - Offset
            #        FibRgFcLcb97      : 0x0382 ( 898bytes )
            #        FibRgFcLcb2000    : 0x03FA ( 1018bytes )
            #        FibRgFcLcb2002    : 0x04DA ( 1242bytes )
            #        FibRgFcLcb2003    : 0x05BA ( 1466bytes )
            #        FibRgFcLcb2007    : 0x0652 ( 1618bytes )
            #    - Size      : 0x0002 ( 2bytes )
            cswNew = Common.BufferControl.ReadWord(pWordDocument, Offset)
            if FibBase["nFib"] == "193" :       # 0x00C1
                if cswNew != 0 :
                    print "\t\t[ERROR] cswNew ( nFib : %s , cswNew : %s )" % (FibBase["nFib"], cswNew) 
                    return OutMapFIB
                
            elif FibBase["nFib"] == "217" or FibBase["nFib"] == "257" and FibBase["nFib"] == "268" :
                if cswNew != 2 :
                    print "\t\t[ERROR] cswNew ( nFib : %s , cswNew : %s )" % (FibBase["nFib"], cswNew)
                    return OutMapFIB
            
            elif FibBase["nFib"] == "274" :     # 0x0112
                if cswNew != 5 :
                    print "\t\t[ERROR] cswNew ( nFib : %s , cswNew : %s )" % (FibBase["nFib"], cswNew)
                    return OutMapFIB
            
            Offset += 2
            
            print "\t   [-] Step 5-2. Mapped FibRgCswNew"
            # Step5-2. Mapped FibRgCswNew
            if cswNew == 0 :
                return OutMapFIB
            
            FibRgCswNew = self.MappedFibRgCswNew(pWordDocument, Offset, cswNew)
            if FibRgCswNew == [] :
                return OutMapFIB
            
            OutMapFIB = MapFIB
            return OutMapFIB
        
        except :
            print traceback.format_exc()
            return OutMapFIB
    
    
    def MappedFibBase(self, pWordDocument):
        try :
            FibBase = {}
            OutFibBase = {}
            
            index = 0
            Position = 0
            
            for index in range( len(szFibBaseList) ) :
                if mFibBase[ index ] == "wIndent" or szFibBaseList[index] == 2 :
                    FibBase[ mFibBase[index] ] = Common.ConvertDataType.unsigned16( Common.BufferControl.ReadWord(pWordDocument, Position) )
                elif szFibBaseList[index] == 4 :
                    FibBase[ mFibBase[index] ] = Common.ConvertDataType.unsigned32( Common.BufferControl.ReadWord(pWordDocument, Position) )
                else :
                    FibBase[ mFibBase[index] ] = Common.BufferControl.Read(pWordDocument, Position, 1)
            
                Position += szFibBaseList[ index ]
        
        except :
            print traceback.format_exc()
            return OutFibBase
        
        OutFibBase = FibBase
        return OutFibBase
    
    
    def MappedFibRgW97(self, pBuf, Offset, szFibRgW97):
        try :
            FibRgW = {}
            OutFibRgW = {}
            
            index = 0
            Position = 0
            
            for index in range( szFibRgW97 / 2 ) :
                # Unsigned Reading...
                FibRgW[ mFibRgW[index] ] = Common.ConvertDataType.unsigned16( Common.BufferControl.ReadWord(pBuf[Offset:], Position) )
                Position += 2
        
        except IndexError :
            print "\t\t[Except] FibRgW97( Index : %x , szFibRgW97 : %s )" % ( index, szFibRgW97 )
            return OutFibRgW
        
        except :
            print traceback.format_exc()
            return OutFibRgW
        
        OutFibRgW = FibRgW
        return OutFibRgW
        
    
    def MappedFibRgLw97(self, pBuf, Offset, szFibRgLw97):
        try :
            FibRglW = {}
            OutFibRglW = {}
            
            index = 0
            Position = 0
            
            for index in range( szFibRgLw97 / 4 ) :
                # Signed Reading...
                FibRglW[ mFibRglW[index] ] = Common.BufferControl.ReadDword(pBuf[Offset:], Position)
                Position += 4
                    
        except :
            print traceback.format_exc()
            return OutFibRglW
    
        OutFibRglW = FibRglW
        return OutFibRglW
    
    
    def MappedFibRgFcLcbBlob(self, pBuf, Offset, nFib):
        try :
            FibRgFcLcb = {}
            OutFibRgFcLcb = {}
                        
            tmpSize = 0
            retList = []
            
            while True :
                # Mapped FibRgFcLcb97
                retList = list( self.MappedFibRgFcLcb97(pBuf, Offset, tmpSize) )
                FibRgFcLcb["97"] = retList[0]

                if nFib == 0x00C1 :
                    break
                
                pBuf = pBuf[retList[1]:]
            
                # Mapped FibRgFcLcb2000
                retList = list( self.MappedFibRgFcLcb2000(pBuf, Offset, tmpSize) )
                FibRgFcLcb["2000"] = retList[0]
                
                if nFib == 0x00D9 :
                    break
                
                pBuf = pBuf[retList[1]:]
                
                # Mapped FibRgFcLcb2002
                retList = list( self.MappedFibRgFcLcb2002(pBuf, Offset, tmpSize) )
                FibRgFcLcb["2002"] = retList[0]
                
                if nFib == 0x0101 :
                    break
                
                pBuf = pBuf[retList[1]:]
                
                # Mapped FibRgFcLcb2003
                retList = list( self.MappedFibRgFcLcb2003(pBuf, Offset, tmpSize) )
                FibRgFcLcb["2003"] = retList[0]
                
                if nFib == 0x010C :
                    break
                
                pBuf = pBuf[retList[1]:]
                
                # Mapped FibRgFcLcb2007
                retList = list( self.MappedFibRgFcLcb2007(pBuf, Offset, tmpSize) )
                FibRgFcLcb["2007"] = retList[0]
                
                break
            
        except :
            print "\t\t[ERROR] FibRgFcLcbBlob.nFib ( %s )" % nFib
            print traceback.format_exc()
            return OutFibRgFcLcb

        OutFibRgFcLcb = FibRgFcLcb        
        return OutFibRgFcLcb
        

    def MappedFibRgFcLcb97(self, pBuf, Offset, Size):
        try :
            FibRgFcLcb97 = {}
            OutFibRgFcLcb97 = {}

            index = 0
            Position = 0
#            print "\t\t" + "=" * 30
#            print "\t\tFibRgFcLcb97"
#            print "\t\t" + "-" * 30
            for index in range( szFibRgFcLcb97 ) :
                FibRgFcLcb97[ mFibRgFcLcb97[index] ] = Common.BufferControl.ReadDword(pBuf[Offset:], Position)
                
#                print "\t\t0x%04X\t%s   : %s" % ( Position, mFibRgFcLcb97[index], FibRgFcLcb97[ mFibRgFcLcb97[index] ] )
                                
                Position += 4           
        
#            print "\t\t" + "-" * 30
            Size = szFibRgFcLcb97 * 4
#            print "\t\t\t Size : %d bytes" % Size
#            print "\t\t" + "=" * 30
        
        except :
            print traceback.format_exc()
            return OutFibRgFcLcb97, 0
        
        OutFibRgFcLcb97 = FibRgFcLcb97
        return OutFibRgFcLcb97, Size


    def MappedFibRgFcLcb2000(self, pBuf, Offset, Size):
        try :
            FibRgFcLcb2000 = {}
            OutFibRgFcLcb2000 = {}

            index = 0
            Position = 0
            for index in range( szFibRgFcLcb2000 ) :
                FibRgFcLcb2000[ mFibRgFcLcb2000[index] ] = Common.BufferControl.ReadDword(pBuf[Offset:], Position)
                Position += 4

            Size = szFibRgFcLcb2000 * 4
            
        except :
            print traceback.format_exc()
            return OutFibRgFcLcb2000, 0
        
        OutFibRgFcLcb2000 = FibRgFcLcb2000
        return OutFibRgFcLcb2000, Size


    def MappedFibRgFcLcb2002(self, pBuf, Offset, Size):
        try :
            FibRgFcLcb2002 = {}
            OutFibRgFcLcb2002 = {}

            index = 0 
            Position = 0
            for index in range( szFibRgFcLcb2002 ) :
                FibRgFcLcb2002[ mFibRgFcLcb2002[index] ] = Common.BufferControl.ReadDword(pBuf[Offset:], Position)
                Position += 4

            Size = szFibRgFcLcb2002 * 4           
        
        except :
            print traceback.format_exc()
            return OutFibRgFcLcb2002, 0
        
        OutFibRgFcLcb2002 = FibRgFcLcb2002
        return OutFibRgFcLcb2002, Size


    def MappedFibRgFcLcb2003(self, pBuf, Offset, Size):
        try :
            FibRgFcLcb2003 = {}
            OutFibRgFcLcb2003 = {}

            index = 0
            Position = 0
            for index in range( szFibRgFcLcb2003 ) :
                FibRgFcLcb2003[ mFibRgFcLcb2003[index] ] = Common.BufferControl.ReadDword(pBuf[Offset:], Position)
                Position += 4

            Size = szFibRgFcLcb2003 * 4
        
        except :
            print traceback.format_exc()
            return OutFibRgFcLcb2003, 0
        
        OutFibRgFcLcb2003 = FibRgFcLcb2003
        return OutFibRgFcLcb2003, Size


    def MappedFibRgFcLcb2007(self, pBuf, Offset, Size):
        try :
            FibRgFcLcb2007 = {}
            OutFibRgFcLcb2007 = {}

            index = 0
            Position = 0
            for index in range( szFibRgFcLcb2007 ) :
                FibRgFcLcb2007[ mFibRgFcLcb2007[index] ] = Common.BufferControl.ReadDword(pBuf[Offset:], Position)
                Position += 4

            Size = szFibRgFcLcb2007 * 4
        
        except :
            print traceback.format_exc()
            return OutFibRgFcLcb2007, 0
        
        OutFibRgFcLcb2007 = FibRgFcLcb2007
        return OutFibRgFcLcb2007, Size

    
    def MappedFibRgCswNew(self, pBuf, Offset, CswNew): 
        try :    
            FibRgCswNew = {}
            OutFibRgCswNew = {}
            
            Position = 0
            nFibNew = Common.BufferControl.ReadWord(pBuf[Offset:], Position)
            if nFibNew != 0x00D9 and nFibNew != 0x0101 and nFibNew != 0x010C and nFibNew != 0x0112 :
                print "\t\t[ERROR] nFibNew ( Offset : 0x%04X, nFibNew : 0x%04X )" % ( Offset, nFibNew )
                return OutFibRgCswNew
            elif nFibNew != 0x0112 :
                FibRgCswNew[ mFibRgCswNewData2000[0] ] = nFibNew
                FibRgCswNew[ mFibRgCswNewData2000[1] ] = Common.BufferControl.ReadWord(pBuf[Offset:], Position+2)
                
#                print "\t\t" + "=" * 30
#                print "\t\tFibRgCswNewData2000"
#                print "\t\t" + "-" * 30
#                
#                print "\t\t0x%04X\t%s   : %s" % ( Offset, mFibRgCswNewData2000[0], FibRgCswNew[ mFibRgCswNewData2000[0] ] )
#                print "\t\t0x%04X\t%s   : %s" % ( Offset+2, mFibRgCswNewData2000[1], FibRgCswNew[ mFibRgCswNewData2000[1] ] )
#                
#                print "\t\t" + "-" * 30
#                print "\t\t Size : %d" % (CswNew * 2) 
#                print "\t\t" + "=" * 30
                
            else : 
                FibRgCswNew[ mFibRgCswNewData2007[0] ] = nFibNew
                FibRgCswNew[ mFibRgCswNewData2007[1] ] = Common.BufferControl.ReadWord(pBuf[Offset:], Position+2)
                FibRgCswNew[ mFibRgCswNewData2007[2] ] = Common.BufferControl.ReadWord(pBuf[Offset:], Position+4)
                FibRgCswNew[ mFibRgCswNewData2007[3] ] = Common.BufferControl.ReadWord(pBuf[Offset:], Position+6)
                FibRgCswNew[ mFibRgCswNewData2007[4] ] = Common.BufferControl.ReadWord(pBuf[Offset:], Position+8)
                
#                print "\t\t" + "=" * 30
#                print "\t\tFibRgCswNewData2007"
#                print "\t\t" + "-" * 30
#                
#                print "\t\t0x%04X\t%s   : %s" % ( Offset, mFibRgCswNewData2000[0], FibRgCswNew[ mFibRgCswNewData2000[0] ] )
#                print "\t\t0x%04X\t%s   : %s" % ( Offset+2, mFibRgCswNewData2000[1], FibRgCswNew[ mFibRgCswNewData2000[1] ] )
#                print "\t\t0x%04X\t%s   : %s" % ( Offset+4, mFibRgCswNewData2000[2], FibRgCswNew[ mFibRgCswNewData2000[2] ] )
#                print "\t\t0x%04X\t%s   : %s" % ( Offset+6, mFibRgCswNewData2000[3], FibRgCswNew[ mFibRgCswNewData2000[3] ] )
#                print "\t\t0x%04X\t%s   : %s" % ( Offset+8, mFibRgCswNewData2000[4], FibRgCswNew[ mFibRgCswNewData2000[4] ] )
#                
#                print "\t\t" + "-" * 30
#                print "\t\t Size : %d" % (CswNew * 2) 
#                print "\t\t" + "=" * 30
                            
        except :
            print traceback.format_exc()
            return OutFibRgCswNew
        
        OutFibRgCswNew = FibRgCswNew
        return OutFibRgCswNew
    
    
    def SumListEntry(self, List):
        try :
            
            sumEntry = 0
            for ListEntry in List :
                sumEntry += ListEntry 
        
            return sumEntry
        
        except :
            print traceback.format_exc()
            return 0
        

#    def GetFIBEntry(self, dictFIB):
#        try :
#            FIBEntry = {}
#            OutFIBEntry = {}
#            
#            entrylist = []
#            for entry in dictFIB :
#                entrylist.append( entry )
#            
#            for n in range( len(entrylist) / 2 ) :
#                if dictFIB[ entrylist[2*n] ] != "0" and dictFIB[ entrylist[2*n+1] ] != "0" :
#                    FIBEntry[ entrylist[2*n] ] = dictFIB[ entrylist[2*n] ]
#                    FIBEntry[ entrylist[2*n+1] ] = dictFIB[ entrylist[2*n+1] ]
#        
#        except :
#            print traceback.format_exc()
#            return OutFIBEntry
#        
#        OutFIBEntry = FIBEntry
#        return OutFIBEntry

        
#----------------------------------------------------------------------------------------------------------------------------------------------------#
#    Word Document ( Office )
#----------------------------------------------------------------------------------------------------------------------------------------------------#

szFibBase = 32
szFibBaseList = [2,2,2,2,2,2,2,4,1,1,2,2,4,4]   # Total Size : 32bytes
#================================================================================================
#                               Offset   : Size : Description
#------------------------------------------------------------------------------------------------
mFibBase = ["wIdent",       # 000 [0x00] : 0x02 : Signature 0xA5EC
            "nFib",         # 002 [0x02] : 0x02 : File Version
            "unused",       # 004 [0x04] : 0x02 : Unused
            "lid",          # 006 [0x06] : 0x02 : Language
            "pnNext",       # 008 [0x08] : 0x02 : AutoText, pnNext * 512
            "Flag",         # 010 [0x0A] : 0x02 : See below...
            "nFibBack",     # 012 [0x0C] : 0x02 : MUST 0x00BF or 0x00C1
            "lKey",         # 014 [0x0E] : 0x04 : See below...
            "envr",         # 018 [0x12] : 0x01 : MUST 0
            "envtFlag",     # 019 [0x13] : 0x01 : See below...
            "reserved3",    # 020 [0x14] : 0x02 : Reserved 
            "reserved4",    # 022 [0x16] : 0x02 : Reserved 
            "reserved5",    # 024 [0x18] : 0x04 : Reserved
            "reserver6"]    # 028 [0x1C] : 0x04 : Reserved
#------------------------------------------------------------------------------------------------
#                             Total Size : 0x20 
#------------------------------------------------------------------------------------------------
#            nFib value    :  Meaning
#            0x00C1        :  if FibBase.Flag ( L ) == True, the LID of the stored style names 
#            0x00D9        :  the LID of the stored style names
#            0x0101        :  the LID of the stored style names
#            0x010C        :  the LID of the stored style names
#            0x0112        :  the LID of the stored style names
#================================================================================================

#szFibRgW = 14       # Total Size : 28bytes
#================================================================================================
#                               Offset   : Size : Description
#------------------------------------------------------------------------------------------------
mFibRgW = ["reserved1",     # 000 [0x00] : 0x02 : Reserved
           "reserved2",     # 002 [0x02] : 0x02 : Reserved
           "reserved3",     # 004 [0x04] : 0x02 : Reserved
           "reserved4",     # 006 [0x06] : 0x02 : Reserved
           "reserved5",     # 008 [0x08] : 0x02 : Reserved
           "reserved6",     # 010 [0x0A] : 0x02 : Reserved
           "reserved7",     # 012 [0x0C] : 0x02 : Reserved
           "reserved8",     # 014 [0x0E] : 0x02 : Reserved
           "reserved9",     # 016 [0x10] : 0x02 : Reserved
           "reserved10",    # 018 [0x12] : 0x02 : Reserved
           "reserved11",    # 020 [0x14] : 0x02 : Reserved
           "reserved12",    # 022 [0x16] : 0x02 : Reserved
           "reserved13",    # 024 [0x18] : 0x02 : Reserved
           "lidFE"]         # 026 [0x1A] : 0x02 : LID whose meaning depends on the nFib Value
#------------------------------------------------------------------------------------------------
#                             Total Size : 0x1C 
#================================================================================================

#szFibRglW = 22      # Total Size : 88bytes
#================================================================================================
#                               Offset   : Size : Description
#------------------------------------------------------------------------------------------------
mFibRglW = ["cbMac",        # 000 [0x00] : 0x04 : the Count of Bytes of those written to the WordDocument stream of file
            "reserved1",    # 004 [0x04] : 0x04 : Reserved
            "reserved2",    # 008 [0x08] : 0x04 : Reserved
            "ccpText",      # 012 [0x0C] : 0x04 : CP Count in Main Document ( signed integer )
            "ccpFtn",       # 016 [0x10] : 0x04 : CP Count in footnote subDocument ( signed integer )
            "ccpHdd",       # 020 [0x14] : 0x04 : CP Count in header subDocument ( signed integer )
            "reserved3",    # 024 [0x18] : 0x04 : Reserved
            "ccpAtn",       # 028 [0x1C] : 0x04 : CP Count in comment subDocument ( signed integer )
            "ccpEdn",       # 032 [0x20] : 0x04 : CP Count in endnote subDocument ( signed integer )
            "ccpTxbx",      # 036 [0x24] : 0x04 : CP Count in textbox subDocument of the Main Document ( signed integer )
            "ccpHdrTxbx",   # 040 [0x28] : 0x04 : CP Count in textbox subDocument of the Header ( signed integer )
            "reserved4",    # 044 [0x2C] : 0x04 : Reserved
            "reserved5",    # 048 [0x30] : 0x04 : Reserved
            "reserved6",    # 052 [0x34] : 0x04 : Reserved
            "reserved7",    # 056 [0x38] : 0x04 : Reserved
            "reserved8",    # 060 [0x3C] : 0x04 : Reserved
            "reserved9",    # 064 [0x40] : 0x04 : Reserved
            "reserved10",   # 068 [0x44] : 0x04 : Reserved
            "reserved11",   # 072 [0x48] : 0x04 : Reserved
            "reserved12",   # 076 [0x4C] : 0x04 : Reserved
            "reserved13",   # 080 [0x50] : 0x04 : Reserved
            "reserved14"]   # 084 [0x54] : 0x04 : Reserved
#------------------------------------------------------------------------------------------------
#                             Total Size : 0x58 
#================================================================================================

szFibRgFcLcb97 = 186  # Total Size : 744bytes
#================================================================================================
#                * Offset ( Table Stream ) Size
#                  - Prefix "fc"     : Location Name
#                  - Prefix "lcb"    : Size Name
#------------------------------------------------------------------------------------------------
mFibRgFcLcb97 = ["fcStshfOrig",         "lcbStshfOrig",         # Ignore
                 "fcStshf",             "lcbStshf",             # struct _STSH              p.472
                 "fcPlcffndRef",        "lcbPlcffndRef",        # struct _PlcffndRef        p.203
                 "fcPlcffndTxt",        "lcbPlcffndTxt",        # struct _PlcffndTxt        p.203
                 "fcPlcfandRef",        "lcbPlcfandRef",        # struct _PlcfandRef        p.196                         
                 "fcPlcfandTxt",        "lcbPlcfandTxt",        # struct _PlcfnadTxt        p.197
                 "fcPlcfSed",           "lcbPlcfSed",           # struct _PlcfSed
                 "fcPlcPad",            "lcbPlcPad",            # Ignore
                 "fcPlcfPhe",           "lcbPlcfPhe",           # struct _Plc               p.27
                 "fcSttbfGlsy",         "lcbSttbfGlsy",         # struct _SttbfGlsy         p.490
                 "fcPlcfGlsy",          "lcbPlcfGlsy",          # struct _PlcfGlsy
                 "fcPlcfHdd",           "lcbPlcfHdd",           # struct _PlcfHdd
                 "fcPlcBteChpx",        "lcbPlcBteChpx",        # struct _PlcBteChpx
                 "fcPlcfBtePapx",       "lcbPlcfBtePapx",       # struct _PlcBtePapx
                 "fcPlcfSea",           "lcbPlcfSea",           # Ignore
                 "fcSttbfFfn",          "lcbSttbfFfn",          # struct _SttbfFfn
                 "fcPlcfFldMom",        "lcbPlcfFldMom",        # struct _PlcFld 
                 "fcPlcfFldHdr",        "lcbPlcfFldHdr",        # struct _PlcFld 
                 "fcPlcfFldFtn",        "lcbPlcfFldFtn",        # struct _PlcFld 
                 "fcPlcfFldAtn",        "lcbPlcfFldAtn",        # struct _PlcFld 
                 "fcPlcfFldMcr",        "lcbPlcfFldMcr",        # Ignore 
                 "fcSttbfBkmk",         "lcbSttbfBkmk",         # struct _SttbfBkmk
                 "fcPlcfBkf",           "lcbPlcfBkf",           # struct _PlcBkf
                 "fcPlcfBkl",           "lcbPlcfBkl",           # struct _PlcfBkl
                 "fcCmds",              "lcbCmds",              # struct _Tcg
                 "fcUnused1",           "lcbUnused1",           # Ignore
                 "fcSttbfMcr",          "lcbSttbfMcr",          # Ignore
                 "fcPrDrvr",            "lcbPDrvr",             # struct _PrDrv
                 "fcPrEnvPort",         "lcbPrEnvPort",         # struct PrEnvPort
                 "fcPrEnvLand",         "lcbPrEnvLand",         # struct _PrEnvLand
                 "fcWss",               "lcbWss",               # struct _Seksf
                 "fcDop",               "lcbDop",               # struct _Dop
                 "fcSttbfAssoc",        "lcbSttbfAssoc",        # struct _SttbfAssoc
                 "fcClx",               "lcbClx",               # struct _Clx
                 "fcPlcfPgdFtn",        "lcbPlcfPgdFtn",        # Ignore
                 "fcAutosaveSource",    "lcbAutosaveSource",    # Ignore
                 "fcGrpXstAtnOwners",   "lcbGrpXstAtnOwners",   # struct _XST
                 "fcSttbfAtnBkmk",      "lcbSttbfAtnBkmk",      # struct _SttbfAtnBkmk
                 "fcUnused2",           "lcbUnused2",           # Ignore
                 "fcUnused3",           "lcbUnused3",           # Ignore
                 "fcPlcSpaMon",         "lcbPlcSpaMon",         # struct _PlcSpa
                 "fcPlcSpaHdr",         "lcbPlcSpaHdr",         # struct _PlcHdr
                 "fcPlcAtnBkf",         "lcbPlcAtnBkf",         # struct _PlcBkf
                 "fcPlcAtnBkl",         "lcbPlcAtnBkl",         # struct _PlcBkl
                 "fcPms",               "lcbPms",               # struct _Pms
                 "fcFormFldSttbs",      "ldbFormFldSttbs",      # Ignore
                 "fcPlcfendRef",        "lcbPlcfendRef",        # struct _PlcfendRef
                 "fcPlcfendTxt",        "lcbPlcfendTxt",        # struct _PlcfendTxt
                 "fcPlcfFldEdn",        "lcbPlcfFldEdn",        # struct _PlcFld 
                 "fcUnused4",           "lcbUnused4",           # Ignore
                 "fcDggInfo",           "lcbDggInfo",           # sturct _OfficeArtContent
                 "fcSttbfRMark",        "lcbSttbfRMark",        # struct _SttbfRMark
                 "fcSttbfCaption",      "lcbSttbfCaption",      # struct _SttbfCaption
                 "fcSttbfAutoCaption",  "lcbSttbfAutoCaption",  # struct _SttbfAutoCaption
                 "fcPlcWkb",            "lcbPlcWkb",            # struct _PlcfWKB
                 "fcPlcfSpl",           "lcbPlcfSpl",           # struct _Plcfspl
                 "fcPlcftxbxTxt",       "lcbPlcftxbxTxt",       # struct _PlcftxbxTxt
                 "fcPlcfFldTxbx",       "lcbPlcfFldTxbx",       # struct _PlcFld 
                 "fcPlcfHdrtxbxTxt",    "lcbPlcfHdrtxbxTxt",    # struct _PlcfHdrtxbxTxt 
                 "fcPlcffldHdrTxdx",    "lcbPlcffldHdrTxdx",    # struct _PlcFld 
                 "fcStwUser",           "lcbStwUser",           # struct _StwUser
                 "fcSttbTtmbd",         "lcbSttbTtmbd",         # struct _SttbTtmbd
                 "fcCookieData",        "lcbCookieData",        # struct _RgCdb
                 "fcPgdMotherOldOld",   "lcbPgdMotherOldOld",   # 
                 "fcBkdMotherOldOld",   "lcbBkdMotherOldOld",   # 
                 "fcPgdFtnOldOld",      "lcbPgdFtnOldOld",      # 
                 "fcBkdFtnOldOld",      "lcbBkdFtnOldOld",      # 
                 "fcPgdEdnOldOld",      "lcbPgdEdnOldOld",      # 
                 "fcBkdEdnOldOld",      "lcbBkdEdnOldOld",      # 
                 "fcSttbfIntlFld",      "lcbSttbfIntlFld",      # Ignore
                 "fcRouteSlip",         "lcbRouteSlip",         # struct _RouteSlip
                 "fcSttbSavedBy",       "lcbSttbSavedBy",       # struct _SttbSavedBy
                 "fcSttbFnm",           "lcbSttbFnm",           # struct _SttbFnm
                 "fcPlfLst",            "lcbPlfLst",            # struct _PlfLst
                 "fcPlfLfo",            "lcbPlfLfo",            # struct _PlfLfo
                 "fcPlcTxbxBkd",        "lcbPlcTxbxBkd",        # struct _PlcftxbxBkd
                 "fcPlcfTxbxHdrBkd",    "lcbPlcfTxbxHdrBkd",    # struct _PlcfTxbxHdrBkd
                 "fcDocUndoWord9",      "lcbDocUndoWord9",      # 
                 "fcRgbUse",            "lcbRgbUse",            # 
                 "fcUsp",               "lcbUsp",               # 
                 "fcUskf",              "lcbUskf",              # 
                 "fcPlcupcRgbUse",      "lcbPlcupcRgbUse",      # struct _Plc
                 "fcPlcupcUsp",         "lcbPlcupcUsp",         # struct _Plc
                 "fcSttbGlsyStyle",     "lcbSttbGlsyStyle",     # struct _SttbGlsyStyle
                 "fcPlgosl",            "lcbPlgosl",            # struct _PlfGosl
                 "fcPlcocx",            "lcbPlcocx",            # struct _RgxOcxInfo
                 "fcPlcfBteLvc",        "lcbPlcfBteLvc",        # numbering field Cache
                 "dwLowDateTime",       "dwHighDateTime",       # struct _FILETIME
                 "fcPlcfLvcPe10",       "lcbPlcfLvcPe10",       # List Level Cache
                 "fcPlcfAsumy",         "lcbPlcfAsumy",         # struct _PlcfAsumy
                 "fcPlcfGram",          "lcbPlcfGram",          # struct _Plcfgram
                 "fcSttbListNames",     "lcbSttbListNames",     # struct _SttbListNames
                 "fcSttbfUssr",         "lcbSttbfUssr"]         # 


szFibRgFcLcb2000 = 30   # Total Size : 120bytes
#================================================================================================
#                Offset ( Table Stream ) Size
#------------------------------------------------------------------------------------------------
mFibRgFcLcb2000 = ["fcPlcfTch",         "lcbPlcfTch",
                   "fcRmdThreading",    "lcbRmdThreading",
                   "fcMid",             "lcbMid",
                   "fcSttbRgtplc",      "lcbSttbRgtplc",
                   "fcMsoEnvelope",     "lcbMsoEnvelope",
                   "fcPlcfLad",         "lcbPlcfLad",
                   "fcRgDofr",          "lcbRgDofr",
                   "fcPlcosl",          "lcbPlcosl",
                   "fcPlcfCookieOld",   "lcbPlcfCookieOld",
                   "fcPgdMotherOld",    "lcbPgdMotherOld",
                   "fcBkdMotherOld",    "lcbBkdMotherOld",
                   "fcPgdFtnOld",       "lcbPgdFtnOld",
                   "fcBkdFtnOld",       "lcbBkdFtnOld",
                   "fcPgdEdnOld",       "lcbPgdEdnOld",
                   "fcBkdEdnOld",       "lcbBkdEdnOld"]


szFibRgFcLcb2002 = 56   # Total Size : 224bytes
#================================================================================================
#                Offset ( Table Stream ) Size
#------------------------------------------------------------------------------------------------
mFibRgFcLcb2002 = ["fcUnused1",         "lcbUnused1",
                   "fcPlcfPgp",         "lcbPlcfPgp",
                   "fcPlcfuim",         "lcbPlcfuim",
                   "fcPlfguidUim",      "lcbPlfguidUim",
                   "fcAtrdExtra",       "lcbAtrdExtra",
                   "fcPlrsid",          "lcbPlrsid",
                   "fcSttbfBkmkFactoid","lcbSttbfBkmkFactoid",
                   "fcPlcfBkfFactoid",  "lcbPlcfBkfFactoid",
                   "fcPlcfcookie",      "lcbPlcfcookie",
                   "fcPlcfBklFactoid",  "lcbPlcfBklFactoid",
                   "fcFactoidData",     "lcbFactoidData",
                   "fcDocUndo",         "lcbDocUndo",
                   "fcSttbfBkmkFcc",    "lcbSttbfBkmkFcc",
                   "fcPlcfBkfFcc",      "lcbPlcfBkfFcc",
                   "fcPlcfBklFcc",      "lcbPlcfBklFcc",
                   "fcSttbfbkmkBPRepairs","lcbSttbfbkmkBPRepairs",
                   "fcPlcfbkfBPRepairs","lcbPlcfbkfBPRepairs",
                   "fcPlcfbklBPRepairs","lcbPlcfbklBPRepairs",
                   "fcPmsNew",          "lcbPmsNew",
                   "fcODSO",            "lcbODSO",
                   "fcPlcfpmiOldXP",    "lcbPlcfpmiOldXP",
                   "fcPlcfpmiNewXP",    "lcbPlcfpmiNewXP",
                   "fcPlcfpmiMixedXP",  "lcbPlcfpmiMixedXP",
                   "fcUnused2",         "lcbUnused2",
                   "fcPlcffactoid",     "lcbPlcffactoid",
                   "fcPlcflvcOldXP",    "lcbPlcflvcOldXP",
                   "fcPlcflvcNewXP",    "lcbPlcflvcNewXP",
                   "fcPlcflvcMixedXP",  "lcbPlcflvcMixedXP"]


szFibRgFcLcb2003 = 56   # Total Size : 224bytes
#================================================================================================
#                Offset ( Table Stream ) Size
#------------------------------------------------------------------------------------------------
mFibRgFcLcb2003 = ["fcHplxsdr",         "lcbHplxsdr",
                   "fcSttbfBkmkSdt",    "lcbSttbfBkmkSdt",
                   "fcPlcfBkfSdt",      "lcbPlcfBkfSdt",
                   "fcPlcfBklSdt",      "lcbPlcfBklSdt",
                   "fcCustomXForm",     "lcbCustomXForm",
                   "fcSttbfBkmkProt",   "lcbSttbfBkmkProt",
                   "fcPlcfBkfProt",     "lcbPlcfBkfProt",
                   "fcPlcfBklProt",     "lcbPlcfBklProt",
                   "fcSttbProtUser",    "lcbSttbProtUser",
                   "fcUnused",          "lcbUnused",
                   "fcPlcfpmiOld",      "lcbPlcfpmiOld",
                   "fcPlcfpmiOldInline","lcbPlcfpmiOldInline",
                   "fcPlcfpmiNew",      "lcbPlcfpmiNew",
                   "fcPlcfpmiNewInline","lcbPlcfpmiNewInline",
                   "fcPlcflvcOld",      "lcbPlcflvcOld",
                   "fcPlcflvcOldInline","lcbPlcflvcOldInline",
                   "fcPlcflvcNew",      "lcbPlcflvcNew",
                   "fcPlcflcxNewInline","lcbPlcflcxNewInline",
                   "fcPgdMother",       "lcbPgdMother",
                   "fcBkdMother",       "lcbBkdMother",
                   "fcAfdMother",       "lcbAfdMother",
                   "fcPgdFtn",          "lcbPgdFtn",
                   "fcBkdFtn",          "lcbBkdFtn",
                   "fcAfdFtn",          "lcbAfdFtn",
                   "fcPgdEdn",          "lcbPgdEdn",
                   "fcBkdEdn",          "lcbBkdEdn",
                   "fcAfdEdn",          "lcbAfdEdn",
                   "fcAfd",             "lcbAfd"]


szFibRgFcLcb2007 = 38   # Total Size : 152bytes
#================================================================================================
#                Offset ( Table Stream ) Size
#------------------------------------------------------------------------------------------------
mFibRgFcLcb2007 = ["fcPlcfmthd",        "lcbPlcfmthd",
                   "fcSttbfBkmkMoveFrom","lcbSttbfBkmkMoveFrom",
                   "fcPlcfBkfMoveFrom", "lcbPlcfBkfMoveFrom",
                   "fcPlcfBklMoveFrom", "lcbPlcfBklMoveFrom",
                   "fcSttbfBkmkMoveTo", "lcbSttbfBkmkMoveTo",
                   "fcPlcfBkfMoveTo",   "lcbPlcfBkfMoveTo",
                   "fcPlcfBklMoveTo",   "lcbPlcfBklMoveTo",
                   "fcUnused1",         "lcbUnused1",
                   "fcUnused2",         "lcbUnused2",
                   "fcUnused3",         "lcbUnused3",
                   "fcSttbfBkmkArto",   "lcbSttbfBkmkArto",
                   "fcPlcfBkfArto",     "lcbPlcfBkfArto",
                   "fcPlcfBklArto",     "lcbPlcfBklArto",
                   "fcArtoData",        "lcbArtoData",
                   "fcUnused4",         "lcbUnused4",
                   "fcUnused5",         "lcbUnused5",
                   "fcUnused6",         "lcbUnused6",
                   "fcOssTheme",        "lcbOssTheme",
                   "fcColorSchemeMapping","lcbColorSchemeMapping"]


#================================================================================================
#    nFibNew        rgCswNewData
#------------------------------------------------------------------------------------------------
#    0x00D9        FibRgCswNewData2000    ( 2Bytes )
#    0x0101        FibRgCswNewData2000    ( 2Bytes )
#    0x010C        FibRgCswNewData2000    ( 2Bytes )
#    0x0112        FibRgCswNewData2007    ( 8Bytes )
#------------------------------------------------------------------------------------------------

mFibRgCswNewData2000 = ["nFibNew",
                        "cQuickSavesNew"]

mFibRgCswNewData2007 = ["nFibNew",
                        "cQuickSavesNew",
                        "lidThemeOther",
                        "lidThemeFE",
                        "lidThemeCS"]
























        
        
        
        
        