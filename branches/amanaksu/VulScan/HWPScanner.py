# -*- coding:utf-8 -*-

# import Public Module
import traceback, binascii, os


# import private module
from ComFunc import BufferControl, FileControl


class HWP():
    @classmethod
    def HWPScan(cls, File):
        try :
            File['logbuf'] += " : HWP"
            Hwp = OperateHWP()           
            
            # Extract Sector Data
            File['logbuf'] += "\n        [Extract] "
            if not Hwp.ExtractSector( File ) : 
                return False
            
            # Trace Sector Data
            File['logbuf'] += "\n        [Trace] "
            
        except :
            print traceback.format_exc()
            return False
        
        return True


    
class OperateHWP():   
    def ExtractSector(self, File):
        try : 
            pBuf = File["pBuf"]
            DirList = File["DirList"]
            
            # Extract Sector By SAT
            SATable = File["SATable"]
            RefSAT = File["RefSAT"]
            for RefEntry in RefSAT :
                for DirEntry in DirList :
                    if RefEntry == DirEntry["EntryName"] :
                        Sector = self.ExtractSectorbySAT(pBuf, SATable, DirEntry["SecID"], DirEntry["szData"])
                        File['logbuf'] += DirEntry["EntryName"] + "\n                  "
                        FileControl.WriteFile("%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"]), Sector)
                        if RefEntry == "RootEntry" :
                            File["RootEntry"] = Sector                        
                        
                        break
                
            # Extract Sector By SSAT
            pBuf = File["RootEntry"]
            if pBuf == "" :
                File['logbuf'] += "[ERROR] RootEntry is NOT Saved"
                return False
            
            SSATable = File["SSATable"]
            RefSSAT = File["RefSSAT"]
            for RefEntry in RefSSAT :
                for DirEntry in DirList :
                    if RefEntry == "" :
                        continue                    
                        
                    if RefEntry == DirEntry["EntryName"] :
                        Sector = self.ExtractSectorbySSAT(pBuf, SSATable, DirEntry["SecID"], DirEntry["szData"])
                        File['logbuf'] += DirEntry["EntryName"] + "\n                  "
                        FileControl.WriteFile("%s_%s_%s.dump" % (File["fname"], DirEntry["EntryName"], DirEntry["szData"]), Sector)
                        break

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
            
                tmpSector += BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x200)
                SecID = Table[SecID]
            
            Sector = BufferControl.Read(tmpSector, 0, Size)
            
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
            
                tmpSector += BufferControl.ReadSectorByBuffer(pBuf, SecID, 0x40)
                SecID = Table[SecID]
            
            Sector = BufferControl.Read(tmpSector, 0, Size)
            
        except :
            print traceback.format_exc()
            return ""
        
        return Sector










