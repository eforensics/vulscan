#!/usr/bin/python
# -*- coding: utf-8 -*-

'''
author: Computists
website: http://computists.tistory.com
last modified: Feb 2013
'''

import wx, os,sys
import wx.lib.mixins.listctrl as listmix
import wx.gizmos as gizmos
import hashlib
import  wx.stc  as  stc

try:
	from agw import flatnotebook as fnb
except ImportError: # if it's not there locally, try the wxPython lib.
	import wx.lib.agw.flatnotebook as fnb

import OLE

MENU_EDIT_DELETE_PAGE = wx.NewId()
OLE_SIGNATURE = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'

### font setting
if wx.Platform == '__WXMSW__':
    face1 = 'Arial'
    face2 = 'Courier New'
    pb = 10
else:
    face1 = 'Helvetica'
    face2 = 'Courier'
    pb = 12

class Anwa(wx.Frame):
    
	def __init__(self, parent):
		super(Anwa, self).__init__(parent,title="MS Word File Analyzer", size=(900,650)) 
		self.CreateMenuBar()
		self.MakeStatusBar()
		self.CreateRightClickMenu()
		self.LayoutItems()
		self.Centre()

	def OnAddPage(self, path):
		if CheckFileFormat(path) == 'OLE' :
			filename = path[path.rfind('\\')+1:]
			self.Freeze()
			
			try:
				self.book.AddPage(self.CreatePage(path), filename)
			except Exception, e:
				print e
				self.log.WriteText("\t[-] Failed to make a page with file information.\n")
				print("Failed to make a page with file information")
			
#			self.book.AddPage(self.CreatePage(path), filename)

			self.Thaw()
		else:
			self.log.WriteText("\t[-] Selected file isn't OLE format\n")
			self.log.WriteText("\t[-] File load failed.\n")

	def CreatePage(self, path):
		ole = OLE.OLE(path, self.log)
		oleData, oleStreamData = ole.OleParsing()
		if oleData == False and oleStreamData == False:
			self.log.WriteText("\t[-] File load failed.\n")
			return False
		else:
			WordData = ole.WordParsing() ### WordData is FIB data
			
			p = middlePanel(self.book, oleData, oleStreamData, WordData, path)
			self.log.WriteText("\t[-] File load done.\n")
				
			return p

	def LayoutItems(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		self.SetSizer(mainSizer)
		spacer1 = wx.Panel(self, -1)
		spacer1.SetBackgroundColour(wx.SystemSettings_GetColour(wx.SYS_COLOUR_3DFACE))
		spacer = wx.Panel(self, -1)
		spacer.SetBackgroundColour(wx.SystemSettings_GetColour(wx.SYS_COLOUR_3DFACE))
		
		middleMainSizer = wx.BoxSizer(wx.HORIZONTAL)

		### book setting ###

		bookStyle = fnb.FNB_NODRAG

		self.book = fnb.FlatNotebook(self, wx.ID_ANY, agwStyle=bookStyle)

		bookStyle &= ~(fnb.FNB_NODRAG)
		bookStyle |= fnb.FNB_ALLOW_FOREIGN_DND 
		self.secondBook = fnb.FlatNotebook(self, wx.ID_ANY, agwStyle=bookStyle)
		self.OnVC8Style()

		# Set right click menu to the notebook
		self.book.SetRightClickMenu(self._rmenu)

		### Layout setup

		### Add middle panel
		mainSizer.Add(self.book, 7, wx.ALL | wx.EXPAND)

		# Add some pages to the second notebook
		mainSizer.Add(self.secondBook, 2, wx.EXPAND)

		self.Freeze()

		self.log = wx.TextCtrl(self.secondBook, -1, "", style=wx.TE_MULTILINE|wx.TE_READONLY)
		self.secondBook.AddPage(self.log, 'Output')
		self.log.WriteText("아흑...TT  \n")

		self.Thaw() 

		mainSizer.Layout()
		self.SendSizeEvent()

	def OnDeletePage(self, event):
		self.book.DeletePage(self.book.GetSelection())

	def CreateRightClickMenu(self):
		self._rmenu = wx.Menu()
		item = wx.MenuItem(self._rmenu, MENU_EDIT_DELETE_PAGE, "Close Tab\tCtrl+F4", "Close Tab")
		self.Bind(wx.EVT_MENU, self.OnDeletePage, item)
		self._rmenu.AppendItem(item)
		
	def MakeStatusBar (self):
		self.statusbar = self.CreateStatusBar(2, wx.ST_SIZEGRIP)
		self.statusbar.SetStatusWidths([-2, -1])
		# statusbar fields
		statusbar_fields = [("Word File Analyzer, Computists @ Jan 2013"),
							("G'Day, Mate!")]

		for i in range(len(statusbar_fields)):
			self.statusbar.SetStatusText(statusbar_fields[i], i)		

	def CreateMenuBar(self):
		menubar = wx.MenuBar()

		### file menu ###
		fileMenu = wx.Menu()

		openmi = wx.MenuItem(fileMenu, wx.ID_OPEN, '&Open')
		fileMenu.AppendItem(openmi)
		self.Bind(wx.EVT_MENU, self.OnButton, openmi)

		fileMenu.AppendSeparator()

		quitmi = wx.MenuItem(fileMenu, wx.ID_EXIT, '&Quit\tCtrl+W')
		fileMenu.AppendItem(quitmi)
		self.Bind(wx.EVT_MENU, self.OnQuit, quitmi)

		menubar.Append(fileMenu, '&File')

		### view menu ###
		viewMenu = wx.Menu()
		menubar.Append(viewMenu, '&View')

		### help menu ###
		helpMenu = wx.Menu()
		item = wx.MenuItem(helpMenu, wx.ID_ANY, "About...", "Shows The About Dialog")
		self.Bind(wx.EVT_MENU, self.OnAboutBox, item)
		helpMenu.AppendItem(item)

		menubar.Append(helpMenu, '&Help')

		self.SetMenuBar(menubar)

	def OnQuit(self, e):
		self.Close()

	def OnButton(self, evt):
		openWildcard = "All files (*.*)|*.*"		

		dlg = wx.FileDialog(
			self, message="Choose a file",
			defaultDir=os.getcwd(), 
			defaultFile="",
			wildcard=openWildcard,
			style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR
			)

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
		if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
			paths = dlg.GetPaths()
			for path in paths:
				# make new pages
				self.log.WriteText('\n[+] Loading %s\n'%path)
				print('\n[+] Loading %s'%path[path.rfind('\\')+1:])
				self.OnAddPage(path)

        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
		dlg.Destroy()

	def OnAboutBox(self, e):
        
		description = "Hello Word(?)"
		licence = "Depend on U!"
		info = wx.AboutDialogInfo()

		info.SetIcon(wx.Icon('mung.png', wx.BITMAP_TYPE_PNG))
		info.SetName('Word File Analyzer')
		info.SetVersion('1.0')
		info.SetDescription(description)
		info.SetCopyright('(C) 2012 - 2013 Computists')
		info.SetWebSite('http://computists.tistory.com/')
		info.SetLicence(licence)
		info.AddDeveloper('Computists@gmail.com')

		wx.AboutBox(info)

	#=== Tab Style ===#

	def OnVC8Style(self):
		style = self.book.GetAGWWindowStyleFlag()

		# remove old tabs style
		mirror = ~(fnb.FNB_VC71 | fnb.FNB_VC8 | fnb.FNB_FANCY_TABS | fnb.FNB_FF2 | fnb.FNB_RIBBON_TABS)
		style &= mirror

		# set new style
		style |= fnb.FNB_VC8

		self.book.SetAGWWindowStyleFlag(style)

#=== List Control ===#

class ListCtrl(wx.ListCtrl, listmix.ListCtrlAutoWidthMixin):
	def __init__(self, parent, ID, pos=wx.DefaultPosition,
				size=wx.DefaultSize, style=0):
		wx.ListCtrl.__init__(self, parent, ID, pos, size, style)
		listmix.ListCtrlAutoWidthMixin.__init__(self)

class middlePanel(wx.Panel):
	def __init__(self, parent, oleData, oleStreamData, WordData, path):
		wx.Panel.__init__(self, parent, -1, style=wx.WANTS_CHARS)
		tID = wx.NewId()
		self.oleData = oleData
		self.oleStreamData = oleStreamData
		self.WordData = WordData

		midSizer = wx.BoxSizer(wx.HORIZONTAL)

		### midTop Setting ###
		midTopSizer = wx.BoxSizer(wx.HORIZONTAL)
		box = wx.StaticBox(self, -1, ">> File Info. ")
		bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)

		t = wx.StaticText(self, -1, "File Path : %s"%path)

		bsizer.Add(t, 0, wx.TOP|wx.LEFT, 5)
		t = wx.StaticText(self, -1, "MD5 : %s"%file_to_md5(path))
		bsizer.Add(t, 00, wx.TOP|wx.LEFT, 5)
		t = wx.StaticText(self, -1, "SHA1 : %s"%file_to_sha1(path))
		bsizer.Add(t, 20, wx.TOP|wx.LEFT, 5)

		### Create buttons at midTop ###
		if self.WordData != None:		

			wordButton = wx.Button(self, -1, "WORD\nStruct\nView", (10,10))
			self.Bind(wx.EVT_BUTTON, self.OnWordButton, wordButton)

			midTopSizer.Add(bsizer, 9, wx.EXPAND|wx.ALL, 8)
			midTopSizer.Add(wordButton, 1, wx.EXPAND|wx.ALL, 8)

		### midLeft Setting ###
		mlBox = wx.StaticBox(self, -1, ">> OLE Dir Tree ")
		mlSizer = wx.StaticBoxSizer(mlBox, wx.VERTICAL)
		tree = DirTree(self, oleData, oleStreamData)
		mlSizer.Add(tree, 1, wx.ALL | wx.EXPAND)

		### midRight Setting ###
		mrBox = wx.StaticBox(self, -1, ">> OLE Dir Info. ")
		mrSizer = wx.StaticBoxSizer(mrBox, wx.VERTICAL)

		self.list = ListCtrl(self, tID, 
								style=wx.LC_REPORT 
								| wx.BORDER_SUNKEN
								#| wx.BORDER_NONE
								| wx.LC_EDIT_LABELS
								#| wx.LC_SORT_ASCENDING
								#| wx.LC_NO_HEADER
								| wx.LC_VRULES
								#| wx.LC_HRULES
								#| wx.LC_SINGLE_SEL
								)
		mrSizer.Add(self.list, 1, wx.EXPAND)

		sizer = wx.BoxSizer(wx.VERTICAL)
		midSizer.Add(mlSizer, 3, wx.ALL | wx.EXPAND, 8)
		midSizer.Add(mrSizer, 7, wx.ALL | wx.EXPAND, 8)
		sizer.Add(midTopSizer, 0, wx.EXPAND)
		sizer.Add(midSizer, 5, wx.EXPAND)
		
		self.PopulateList()

		self.SetSizer(sizer)
		self.SetAutoLayout(True)

		# for wxMSW
		self.list.Bind(wx.EVT_COMMAND_RIGHT_CLICK, self.OnRightClick)

		# for wxGTK
		self.list.Bind(wx.EVT_RIGHT_UP, self.OnRightClick)

	def PopulateList(self):
		self.list.InsertColumn(0, "Num", wx.LIST_FORMAT_CENTER)
		self.list.InsertColumn(1, "Directory name")
		self.list.InsertColumn(2, "Type", wx.LIST_FORMAT_CENTER)
		self.list.InsertColumn(3, "Left Child", wx.LIST_FORMAT_CENTER)
		self.list.InsertColumn(4, "Right Child", wx.LIST_FORMAT_CENTER)
		self.list.InsertColumn(5, "Root Child", wx.LIST_FORMAT_CENTER)
		self.list.InsertColumn(6, "SectionID", wx.LIST_FORMAT_CENTER)
		self.list.InsertColumn(7, "Size")
		self.list.SetColumnWidth(0, 40)
		self.list.SetColumnWidth(1, 130)
		self.list.SetColumnWidth(2, 50)
		self.list.SetColumnWidth(3, 80)
		self.list.SetColumnWidth(4, 80)
		self.list.SetColumnWidth(5, 80)
		self.list.SetColumnWidth(6, 70)
		self.list.SetColumnWidth(7, 50)
#		self.list.SetColumnWidth(7, wx.LIST_AUTOSIZE)

		items = self.oleData.items()
		for key, data in items:
#			print(key, data)
			for i in range(len(data)):
				if i == 0:				
					index = self.list.InsertStringItem(key, data[i]) 
				else:
					try: ### 02e8eb120b5caa73264751b85131d21a.bin
						self.list.SetStringItem(index, i, data[i])
					except:
						self.list.SetStringItem(index, i, StrToHexStr(data[i]))

        # show how to change the colour of a couple items
		for key, data in items:
			if key%3 == 0:
				item = self.list.GetItem(key)
				item.SetTextColour(wx.BLACK)
				self.list.SetItem(item)
			if key%3 == 1:
				item = self.list.GetItem(key)
				item.SetTextColour(wx.RED)
#				item.SetBackgroundColour([176,196,222])
				self.list.SetItem(item)
			if key%3 == 2:
				item = self.list.GetItem(key)
				item.SetTextColour(wx.BLUE)
				self.list.SetItem(item)

		self.currentItem = 0

	def OnWordButton(self, evt): ### using OLE data made in Middle Panel
		win = WordView(self, -1, "Word File Struct View", self.WordData, 
							size=(450, 700),
							style = wx.DEFAULT_FRAME_STYLE)
		win.Show(True)

	def OnRightClick(self, event):
		print "OnRightClick %s\n" % self.list.GetItemText(self.currentItem)

		# only do this part the first time so the events are only bound once
		if not hasattr(self, "popupID1"):
			self.popupID1 = wx.NewId()
			self.popupID2 = wx.NewId()

			self.Bind(wx.EVT_MENU, self.OnPopupOne, id=self.popupID1)
			self.Bind(wx.EVT_MENU, self.FileSave, id=self.popupID2)

		# make a menu
		menu = wx.Menu()
		# add some items
		menu.Append(self.popupID1, "Hex View")
		menu.Append(self.popupID2, "Stream Dump")

		# Popup the menu.  If an item is selected then its handler
		# will be called before PopupMenu returns.
		self.PopupMenu(menu)
		menu.Destroy()

	def OnPopupOne(self, event):
		index = self.list.GetFirstSelected()
		itemNum = self.list.GetItemText(index)
		itemName = self.getColumnText(index, 1)
		print(itemNum, itemName)
		win = HexViewer(self, -1, "Simple Hex Viewer", itemName, self.oleStreamData[int(itemNum)], 
							size=(700, 500),
							style = wx.DEFAULT_FRAME_STYLE)
		win.Show(True)

	def FileSave(self, event):
		index = self.list.GetFirstSelected()
		itemNum = self.list.GetItemText(index)
		itemName = self.getColumnText(index, 1)

		dumpWildcard = "All files (*.*)|*.*|"\
					"Data files (*.dat)|*.dat|" \
					"Bin files (*.bin)|*.bin"

		dlg = wx.FileDialog(
			self, message="Save file as ...", defaultDir=os.getcwd(), 
			defaultFile=itemName, wildcard=dumpWildcard, style=wx.SAVE
			)
		dlg.SetFilterIndex(2)

		if dlg.ShowModal() == wx.ID_OK:
			path = dlg.GetPath()
		f = open(path, 'wb')
		f.write(self.oleStreamData[int(itemNum)])
		f.close()

		dlg.Destroy()

	def getColumnText(self, index, col):
		item = self.list.GetItem(index, col)
		return item.GetText()


class MySTC(stc.StyledTextCtrl):
    def __init__(self, parent, ID):
        stc.StyledTextCtrl.__init__(self, parent, ID)

        self.Bind(wx.EVT_WINDOW_DESTROY, self.OnDestroy)

    def OnDestroy(self, evt):
        # This is how the clipboard contents can be preserved after
        # the app has exited.
        wx.TheClipboard.Flush()
        evt.Skip()


class HexViewer(wx.Frame):
	def __init__(self, *args, **kw):
		itemName = args[3] ### separate ole data in args
		oleStreamData = args[4]
		args = args[:3]

		wx.Frame.__init__(self, *args, **kw)

		p = wx.Panel(self, -1, style=wx.NO_FULL_REPAINT_ON_RESIZE)
		ed = MySTC(p, -1)
		s = wx.BoxSizer(wx.HORIZONTAL)
		s.Add(ed, 1, wx.EXPAND)
		p.SetSizer(s)
		p.SetAutoLayout(True)

		demoText = self.MakeHexView(oleStreamData)
		ed.SetText(demoText)
		ed.SetReadOnly(True)

		# make some styles
		ed.StyleSetSpec(stc.STC_STYLE_DEFAULT, "size:%d,face:%s" % (pb, face2))
		ed.StyleClearAll()

		# line numbers in the margin
		ed.SetMarginType(0, stc.STC_MARGIN_NUMBER)
		ed.SetMarginWidth(0, 22)
		ed.StyleSetSpec(stc.STC_STYLE_LINENUMBER, "size:%d,face:%s" % (pb-2, face1))

	### Ref> http://mwultong.blogspot.com/2007/04/python-hex-viewer-file-dumper.html
	def MakeHexView(self, oleStreamData):
		offset = 0
		retData = []
		while True:
			output = []
			buf16 = oleStreamData[offset:offset+16] # 파일을 16바이트씩 읽어 버퍼에 저장
			buf16Len = len(buf16) # 버퍼의 실제 크기 알아내기
			if buf16Len == 0: break

			output.append("%08X:  " % offset) # Offset(번지)을, 출력 버퍼에 쓰기

			for i in range(buf16Len): # 헥사 부분의 헥사 값 16개 출력 (8개씩 2부분으로)
				if (i == 8): output.append(" ") # 8개씩 분리
				output.append("%02X " % (ord(buf16[i]))) # 헥사 값 출력

			for i in range( ((16 - buf16Len) * 3) + 1 ): # 한 줄이 16 바이트가 되지 않을 때, 헥사 부분과 문자 부분 사이에 공백들 삽입
				output.append(" ")
			if (buf16Len < 9):
				output.append(" ") # 한줄이 9바이트보다 적을 때는 한칸 더 삽입

			for i in range(buf16Len): # 문자 구역 출력
				if (ord(buf16[i]) >= 0x20 and ord(buf16[i]) <= 0x7E): # 특수 문자 아니면 그대로 출력
				  output.append(buf16[i])
				else: output.append(".") # 특수문자, 그래픽문자 등은 마침표로 출력
			output.append('\n')

			offset += 16 # 번지 값을 16 증가
			result = "".join(output)
			retData.append(result)
		result = "".join(retData)		
		return result

class WordView(wx.Frame):
    def __init__(self, *args, **kw):
		wordData = args[3] ### separate word data in args
		args = args[:3]

		wx.Frame.__init__(self, *args, **kw)

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		tree = WordPanel(self, wordData)
		mainSizer.Add(tree, 1, wx.ALL | wx.EXPAND)
				
		self.SetSizer(mainSizer)

class MyTreeCtrl(wx.TreeCtrl):
    def __init__(self, parent, id, pos, size, style):
        wx.TreeCtrl.__init__(self, parent, id, pos, size, style)

class DirTree(wx.Panel):
	def __init__(self, parent, oleData, oleStreamData):
		wx.Panel.__init__(self, parent, -1, style=wx.WANTS_CHARS)
		self.Bind(wx.EVT_SIZE, self.OnSize)
		tID = wx.NewId()

		self.tree = MyTreeCtrl(self, tID, wx.DefaultPosition, wx.DefaultSize,
								wx.TR_DEFAULT_STYLE
								#wx.TR_HAS_BUTTONS
								#| wx.TR_EDIT_LABELS
								#| wx.TR_MULTIPLE
								#| wx.TR_HIDE_ROOT
								)

		rootNum = 0
		self.oleData = oleData.items()
		for i,data in self.oleData: ### find the Root Entry
			if data[1].upper() == 'ROOT ENTRY' or data[2] == 5 or data[2] == 1:
				rootNum = int(i)

				root = self.tree.AddRoot(self.oleData[rootNum][1][1]) ### make the Root Entry / Name == Root Entry | type = 1, 5
				self.tree.SetPyData(root, None)

				self.MakeTree(root, int(self.oleData[rootNum][1][5]))

	def MakeTree (self, root, treeNum):
		treeName = self.oleData[treeNum][1][1] ### get a name of root dor
		try: ### 02e8eb120b5caa73264751b85131d21a.bin
			child = self.tree.AppendItem(root, treeName) ### make child tree
		except:
			child = self.tree.AppendItem(root, StrToHexStr(treeName)) ### make child tree
		self.tree.SetPyData(child, None)
		self.tree.Expand(root)

		if self.oleData[treeNum][1][5] != 'None':
			newNum = int(self.oleData[treeNum][1][5]) ### get a number of root dir
			self.MakeTree(child, newNum)
		if self.oleData[treeNum][1][3] != 'None':
			newNum = int(self.oleData[treeNum][1][3]) ### get a number of left child dir
			self.MakeTree(root, newNum)
		if self.oleData[treeNum][1][4] != 'None':
			newNum = int(self.oleData[treeNum][1][4]) ### get a number of right child dir
			self.MakeTree(root, newNum)
		return
			
	def OnSize(self, event):
		w,h = self.GetClientSizeTuple()
		self.tree.SetDimensions(0, 0, w, h)

class WordPanel(wx.Panel):
	def __init__(self, parent, wordData):
		wx.Panel.__init__(self, parent, -1)
		self.Bind(wx.EVT_SIZE, self.OnSize)
		self.backColor = [230,245,255]
#		self.backColor = [246,222,246]

		self.tree = gizmos.TreeListCtrl(self, -1, style =
										wx.TR_DEFAULT_STYLE
										#| wx.TR_HAS_BUTTONS
										#| wx.TR_TWIST_BUTTONS
										#| wx.TR_ROW_LINES
										| wx.TR_COLUMN_LINES
										#| wx.TR_NO_LINES 
										| wx.TR_FULL_ROW_HIGHLIGHT
										)

		# create some columns
		self.tree.AddColumn("Structure")
		self.tree.AddColumn("Value")
		self.tree.SetMainColumn(0) # the one with the tree in it...
		self.tree.SetColumnWidth(0, 175)
		self.tree.SetColumnWidth(1, 300)

		self.root = self.tree.AddRoot("FIB structure")
		self.tree.SetItemText(self.root, "The File Information Block", 1)
		self.tree.SetItemBackgroundColour(self.root,self.backColor)

		FIBnames = wordData._fields
		FIBvalues = wordData.__getnewargs__()

		self.MakeWordTree(self.root, FIBnames, FIBvalues)

		self.tree.Expand(self.root)

		self.Bind(wx.EVT_TREE_ITEM_EXPANDED, self.OnItemExpanded, self.tree)
		self.Bind(wx.EVT_TREE_ITEM_COLLAPSED, self.OnItemCollapsed, self.tree)

		self.MakeItemBackgroundColor()
	
	def OnItemExpanded (self, event):
		self.MakeItemBackgroundColor()

	def OnItemCollapsed (self, event):
		self.MakeItemBackgroundColor()

	def MakeItemBackgroundColor(self):
		visItem = self.tree.GetFirstExpandedItem()

		xor = 0
		while(True):
			visItem = self.tree.GetNextExpanded(visItem)
			if visItem.IsOk():			
				xor = xor^1
				if xor:
					self.tree.SetItemBackgroundColour(visItem,self.backColor)
				else:
					self.tree.SetItemBackgroundColour(visItem,[255,255,255])					
			else:
				break
		return

	def MakeWordTree (self, root, names, values):
		for i in range(len(names)) :
			child = self.tree.AppendItem(root, names[i])
			if isinstance(values[i],tuple):
				childNames = values[i]._fields
				childValues = values[i].__getnewargs__()
				self.MakeWordTree(child, childNames, childValues)
			else:
				value = StrToHexStr(values[i])
				self.tree.SetItemText(child, value, 1)
		return

	def OnSize(self, evt):
		self.tree.SetSize(self.GetSize())
		return

def StrToHexStr (String):
	try:
		hexData = [ "\\x%02X"%ord(data) for data in String ]
		result = "".join(hexData)
	except:
		if isinstance(String, int) and String != 0:		
			result = '%d [%02Xh]'%(String, String)
		else:
			result = str(String)
	return result

def CheckFileFormat(path):
	sig = open(path,'rb').read(8)
	if sig == OLE_SIGNATURE:
		return 'OLE'
	else:
		return False

### Ref> http://pythonadventures.wordpress.com/2011/11/17/md5-hash-of-a-text-file-crack-an-md5-hash/
def file_to_md5(filename): 
	md5 = hashlib.md5() 
	with open(filename,'rb') as f: 
		for chunk in iter(lambda: f.read(8192), ''): 
			md5.update(chunk) 
	return md5.hexdigest() 

def file_to_sha1(filename): 
	md5 = hashlib.sha1() 
	with open(filename,'rb') as f: 
		for chunk in iter(lambda: f.read(8192), ''): 
			md5.update(chunk) 
	return md5.hexdigest() 

def Main():    
	ex = wx.App(False)
	WFA = Anwa(None)
	WFA.Show(True)
	ex.MainLoop()

	return

if __name__ == '__main__':
	Main()
