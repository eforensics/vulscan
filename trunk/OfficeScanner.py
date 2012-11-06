# -*- coding:utf-8 -*-

# import Public Module
import traceback, sys, os

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
        
        