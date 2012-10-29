# -*- coding:utf-8 -*-

# import Public Module
import optparse, traceback, sys, os, ftplib, shutil

# import Private Module
import Common



class Office():
    @classmethod
    def Scan(cls, File):
        
        try :
            
            return True
        
        except :
            print traceback.format_exc()
            return False
        
        