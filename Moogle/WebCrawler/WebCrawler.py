
# External Import Module
from traceback import format_exc
from optparse import OptionParser
from bs4 import BeautifulSoup
import robotparser
import urllib2
import time
import os, sys
import sqlite3
import re
import binascii


# Internal Import Module
from Common import CFile, CBuffer


class CDB():
    def __init__(self, WEB_CRAWLER_DB_NAME):
        #    struct _DB
        #    =======================================
        #    IDN001 | IDN002 | IDN003  | IDN004
        #    ---------------------------------------
        #    Cate   |   url  |  state  | descriptor
        #    =======================================
        self.conn = sqlite3.connect(WEB_CRAWLER_DB_NAME)
        self.cursor = self.conn.cursor()
        self.cursor.execute("CREATE TABLE IF NOT EXISTS urls(url text, state int, descriptor text)")
        self.cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS IDN001 ON urls(url)")
        self.cursor.execute("CREATE INDEX IF NOT EXISTS IDN002 ON urls(state)")
        self.cursor.execute("CREATE INDEX IF NOT EXISTS IDN003 ON urls(descriptor)")
    def __del__(self):
        self.conn.commit()
        self.cursor.close()        
    def fnInsertDB(self, s_URL, s_Text, n_State=0):
        # [IN]    s_URL
        # [IN]    n_State
        # [OUT]
        #    - [SUCCESS] 1
        #    - [ERROR] 0
        try :
            
            self.cursor.execute("INSERT INTO urls VALUES (%s, %d, %s)" % (s_URL, n_State, s_Text))
            
        except :
            print format_exc()
            return 0
        
        return 1
    def fnSelectUnCrawledURL(self):
        try :
            
            self.cursor.execute("SELECT * FROM urls where state = 0")
            return [ row[0] for row in self.cursor.fetchall() ]
            
        except :
            print format_exc()
    def fnUpdateURL(self, s_URL, n_State=1):
        # [IN]    s_URL
        # [IN]    n_State
        # [OUT]
        try :
            
            self.cursor.execute("UPDATE urls SET state=%d WHERE url=%s" % (n_State, s_URL))
            
        except :
            print format_exc()
    def fnIsCrawledURL(self, s_URL):
        # [IN]    s_URL
        # [OUT]
        try :
            
            self.cursor.execute("SELECT COUNT(*) FROM urls WHERE url=%s and state=1" % s_URL)
            l_Ret = self.cursor.fetchone()
            
        except :
            print format_exc()
            
        return l_Ret[0]

class CUTIL():
    def __del__(self):
        return
    def fnGetFilePath(self, l_path, s_fname, s_ext):
        # Create By Hanul93
        # - Referred : https://www.evernote.com/shard/s24/sh/66bc80ea-d1a1-449e-b504-fc6e4e0fc3f2/89cd44a81619173485e0bae4a603a58d
        # [IN]    l_path
        # [IN]    s_fname
        # [IN]    s_ext
        # [OUT]   s_fpath
        #    - [SUCCESS] s_fpath : path string
        #    - [FAILURE] s_fpath : ""
        s_fpath = ""
        
        try :
            
            for s_path in l_path :
                s_tmpfpath = "%s\\%s%s" % (s_path, s_fname, s_ext)
                
                if os.path.exists( s_tmpfpath ) == True :
                    s_fpath = s_tmpfpath  
                    break
            
        except :
            print format_exc()
            return None
            
        return s_fpath
    def fnGetFilePathEx(self, s_fname):
        # Create By Hanul93
        # [IN]    s_fname
        # [OUT]   s_fpath
        #    - [SUCCESS] s_fpath : path string
        #    - [FAILURE] s_fpath : None
        #    - [FAILURE] s_fpath : ""
        s_fpath = ""
        
        try :
            
            if s_fname[1] != ":" :
                s_path = os.environ["path"]
                l_path = s_path.split(";")
                
                s_fpath = self.fnGetFilePath(l_path, s_fname, "")
                if s_fpath != None :
                    return s_fpath
                
                s_fpath = self.fnGetFilePath(l_path, s_fname, ".txt")
                if s_fpath != None :
                    return s_fpath
            
            else :
                if os.path.exists( s_fname ) == True :
                    s_fpath = s_fname
                    return s_fpath
            
        except :
            print format_exc()
            return None
        
        return s_fpath

class CRobots():
    def fnCheckRobots(self, s_CrawlerName, s_URL):
        # [IN]    s_URL
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        try :
            
            p_Robot = robotparser.RobotFileParser(s_URL + "robots.txt")
            p_Robot.read()
            return p_Robot.can_fetch(s_CrawlerName, s_URL)

        except IOError :
            return None

        except :
            print format_exc()
            return None 
    def fnGetContents(self, s_UserAgent, s_URL, n_delay=1):
        # [IN]    s_URL
        # [IN]    n_delay
        # [OUT]   s_Contents
        #    - [SUCCESS] s_Contents : String Data
        #    - [FAILURE] s_Contents : ""
        #    - [ERROR] s_Contents : None
        s_Contents = ""
        try :
            
            time.sleep( n_delay )
            
            p_Opener = urllib2.build_opener()
            p_Opener.addheaders = [("User-agent", s_UserAgent)]
            s_Contents = p_Opener.open(s_URL).read()
            
        except urllib2.HTTPError :
            return None
        
        except AttributeError :
            print ""
            print s_URL
            print format_exc()
            return None
            
        except :
            print format_exc()
            return None
        
        return s_Contents
#    def fnGetHREF(self, Soup):
#        # [IN]    Soup
#        # [OUT]   l_HREF
#        #    - [SUCCESS] l_HREF : String Data ( [0] : Title, [1] : URL )
#        #    - [FAILURE] l_HREF : []
#        #    - [ERROR] l_HREF : None
#        dl_HREF = []
#        l_HREF_Title = []
#        l_HREF_URLs = []
#        
#        try :
#        
#            l_tmpHREF = Soup('a')
#            if l_tmpHREF == [] :
#                return dl_HREF
#            
#            for s_tmpHREF in l_tmpHREF :  
#                if s_tmpHREF.get("title") :  
#                    l_HREF_Title.append( s_tmpHREF.get("title") )           # type( s_tmpHREF.get("title") ) : Unicode
#                    l_HREF_URLs.append( s_tmpHREF.get("href").split("?")[0] )  # type( s_tmpHREF.get("href").split("?") ) : List
#        
#            dl_HREF.append( l_HREF_Title )
#            dl_HREF.append( l_HREF_URLs )
#        
#        except UnicodeEncodeError :
#            print format_exc()
#        
#        except :
#            print format_exc()
#            return None
#        
#        return dl_HREF
#    def fnGetFrames(self, Soup):
#        # [IN] Soup
#        # [OUT] l_Frames
#        #    - [SUCCESS] l_Frames : String Data ( [0] : Title, [1] : URL )
#        #    - [FAILURE] l_Frames : ""
#        #    - [ERROR] l_Frames : None
#        l_Frames = []
#        l_Frames_Title = []
#        l_Frames_URLs = []
#        
#        try :
#            
#            l_tmpFrames_URLs = Soup("frame")
#            if l_tmpFrames_URLs == [] :
#                return l_Frames_URLs
#            
#            for s_tmpFrames_URLs in l_tmpFrames_URLs :
#                if s_tmpFrames_URLs.get( "title" ) : 
#                    l_Frames_Title.append( s_tmpFrames_URLs.get( "title" ) )
#                    l_Frames_URLs.append( s_tmpFrames_URLs.get("src") )
#            
#            l_Frames.append( l_Frames_Title )
#            l_Frames.append( l_Frames_URLs )
#            
#        except :
#            print format_exc()
#            return None
#        
#        return l_Frames
    def fnGetArticle(self, Soup):
        # [IN]    Soup
        # [OUT]   l_Article
        #    - [SUCCESS] l_Article : Article Data
        #    - [FAILURE] l_Article : []
        #    - [ERROR] l_Article : None
        dl_Article = []
        l_Text = []
        l_URLs = []
        l_Http = []
        try : 
            
            l_Http = Soup("a")
            if l_Http == [] :
                return dl_Article
            
            for s_Http in l_Http :
                if s_Http.get("href") :
                    l_URLs.append( s_Http.get("href") )
                    
                    # Case 1. Exist "Title"
                    if s_Http.get("title") :
                        l_Text.append( s_Http.get("title") )
                    # Case 2. Exist "get_txt"
                    else :
                        l_Text.append( s_Http.get_text() )
                
            dl_Article.append( l_Text )
            dl_Article.append( l_URLs )
        
        except :
            print format_exc()
            return None
        
        return dl_Article
    def fnGetCrawling(self, s_UnCrawled):
        # [IN]    s_UnCrawled
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        try :
            
            return True
            
        except :
            print format_exc()
            return None
        
        return True


class CPrintLog():
    def __del__(self):
        return        
    def fnPrintData(self, l_Set):
        # [IN]    l_Set ( l_Set[0] : Title, l_Set[1] : URL )
        # [OUT]   
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        try :
            
            print "=" * 80
            for n_Index in range( str(l_Set[0]).split(",").__len__() ) :
                print str(l_Set[0]).split(",")[n_Index],
                print "|",
                print str(l_Set[1]).split(",")[n_Index]
                print "-" * 80
            print "=" * 80
        
        except :
            print format_exc()
            return None
        
        return True

class CInit():
    def __init__(self):
        print "-----------------------------------------------------------------"
        print "  WebCrawler ver 1.0                 Copyright (c) 2013, Moogle"
        print "-----------------------------------------------------------------"
        print ""
    def fnGetOptions(self):
        try :
            # Default Option
            parser = OptionParser(usage="%prog -d <Folder>", version="%prog 1.0")
            parser.add_option("-d", "--dir", metavar="DIR", type="string", help="config file directory")
            parser.add_option("-f", "--file", metavar="FILE", type="string", help="config file full path")
            (Options, args) = parser.parse_args()
                
            # Except Check Parameter
            if sys.argv.__len__() < 2 :
                parser.print_help()
                return None
        
        except :
            print format_exc()
            return None
            
        return Options
    def fnGetEnvInfo(self, Main, i_Options):
        # [IN]    Main
        # [IN]    i_Options
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        try :
            s_config_path = ""
            
            if i_Options.dir :
                s_config_path = "%s\\config.ini" % i_Options.dir 
            elif i_Options.file :
                s_config_path = i_Options.file
            else :
                print "[-] Error - fnGetEnvInfoEx() : Check Parameter"
                return False
            
            if not os.path.exists( s_config_path ) :
                print "[-] Error - Do not EXIST %s" % s_config_path
                return False
            
            # Set Default Information
            l_config = open(s_config_path, "rb").read().split("\r\n")
            for s_config in l_config :
                s_header = s_config.split("=")[0]
                s_body = s_config.split("=")[1]
                if s_header == "URLPATH" :
                    Main.URL_PATH = s_body
                elif s_header == "CATEGORY" :
                    Main.CATEGORY_FILE = s_body
                elif s_header == "USER_AGENT_NAME" :
                    Main.USER_AGENT_NAME = s_body
                elif s_header == "WEB_CRAWLER_DB_NAME" :
                    Main.WEB_CRAWLER_DB_NAME = s_body
                else :
                    print "[-] Error - Default Information Setting"
                    return False
            
        except :
            print format_exc()
            return None
        
        return True
    def fnGetCategoryInfo(self, Main):
        # [IN]    Main
        # [OUT]
        #    - [SUCCESS] True
        #    - [FAILURE] False
        #    - [ERROR] None
        try :
            s_category_path = "%s\\%s" % (Main.URL_PATH, Main.CATEGORY_FILE)
            if not os.path.exists( s_category_path ) :
                print "[-] Error - Do not EXIST %" % s_category_path
                return False
            
            # Setting Default Information
            l_category = open(s_category_path, "rb").read().split("\r\n")
            
            # Check File Exists
            for s_category in l_category :
                if os.path.exists( "%s\\%s" % (Main.URL_PATH, s_category) ) :
                    Main.l_Category.append( s_category )
            
        except :
            print format_exc()
            return None
        
        return True
    def fnCheckRobots(self, Main, s_Category):
        # [IN]    Main
        # [IN]    s_Category
        # [OUT]   l_Robots
        #    - [SUCCESS] URL List
        #    - [FAILURE] 
        #    - [ERROR] None
        try :
            
            l_Robots = []
            Robots = CRobots()
            l_URLs = open("%s\\%s" % (Main.URL_PATH, s_Category), "rb").read().split("\r\n")
            for s_URL in l_URLs :
                if Robots.fnCheckRobots(Main.USER_AGENT_NAME, s_URL) :
                    l_Robots.append( s_URL )
            
        except :
            print format_exc()
            return None
        
        return l_Robots
        

class CMain():
    URL_PATH = ""                   # Set fnGetEnvInfo()
    CATEGORY_FILE = ""              # Set fnGetEnvInfo()
    USER_AGENT_NAME = ""            # Set fnGetEnvInfo()
    WEB_CRAWLER_DB_NAME = ""        # Set fnGetEnvInfo()
    
    l_Category = []                 # Set fnGetCategoryInfo()
    dl_Robots = []                  # Set 
    def Main(self):
        Init = CInit()
        
        print "[*] Init"
        i_Options = Init.fnGetOptions()
        if i_Options == None :
            print "[-] Error :: fnGetOptions()"
            return False
        
        # Get Environment Information
        print "    - Get Env Info..........",
        b_Ret = Init.fnGetEnvInfo(self, i_Options)
        if b_Ret == False :
            print "Failure :: fnGetEnvInfo()"
            return False
        elif b_Ret == None :
            print "Error :: fnGetEnvInfo()"
            return False
        else :
            print "OK!"
        
        # Get Category Information
        print "    - Get Category Info.....",
        b_Ret = Init.fnGetCategoryInfo(self)
        if b_Ret == False :
            print "Failure :: fnGetCategoryInfo()"
            return False
        elif b_Ret == None :
            print "Error :: fnGetCategoryInfo()"
            return False
        elif self.l_Category == [] :
            print "Do not Exist Category File List"
            return True
        else :
            print "OK!"
        
        # Check Robots
        print "    - Check Robots..........",
        for s_Category in self.l_Category :
            l_Robots = Init.fnCheckRobots(Main, s_Category)
            if l_Robots == None :
                print "Error :: fnCheckRobots()"
                return False
            self.dl_Robots.append( l_Robots )
        if self.dl_Robots == [] :
            print "Do not exist enable URLs"
            return True
        else :
            print "OK!"
        
        # Grab Information
#            print "%s %s : %s" % (l_Robots[0], self.l_Category[ l_Robots[0] ].split(".")[0], str(l_Robots[1]) )
        print "[*] DataBase"
        Robots = CRobots()
        DB = CDB()
        for l_Robots in enumerate( self.dl_Robots ) :
            if l_Robots[1] == [] :
                continue
            
            print "    - %s.................." % self.l_Category[ l_Robots[0] ].split(".")[0],
            for s_Robot in l_Robots[1] :            
                s_Contents = Robots.fnGetContents(self.USER_AGENT_NAME, s_Robot)
                if s_Contents == "" :
                    print "Failure :: fnGetContents()"
                    break
                elif s_Contents == None :
                    print "Error :: fnGetContents()"
                    break
                                
                dl_Article = Robots.fnGetArticle( BeautifulSoup(s_Contents) )
                if dl_Article == [] :
                    print "Failure :: fnGetArticle()"
                    print dl_Article
                    break
                elif dl_Article == None :
                    print "Error :: fnGetArticle()"
                    break
            
#                for s_Article in dl_Article :
#                    DB.fnInsertDB(s_URL, s_Text)
            
            print "OK!"


if __name__ == "__main__" :
    Main = CMain()
    Main.Main()
    exit(0)

#        # Insert DB
#        n_Success = 0
#        for s_Article in dl_Article :
#            s_URL = ""
#            s_Text = ""
#            n_Success += DB.fnInsertDB(s_URL, s_Text)
#        print "Inserted %d New Pages." % n_Success
#
#
#    # Crawling URLs
#    print "[*] Crawling"
#    while True :
#        for s_UnCrawled in DB.fnSelectUnCrawledURL() :
#            try :
#                b_Ret = Robots.fnGetCrawling( s_UnCrawled )
#                if b_Ret == False :
#                    print "[-] Failure - fnGetCrawling( %s )" % s_UnCrawled
#                    continue
#                elif b_Ret == None :
#                    print "[-] Error - fnGetCrawling( %s )" % s_UnCrawled
#                    continue
#                
#            except :
#                print format_exc()
#                DB.fnUpdateURL(s_UnCrawled, -1)
