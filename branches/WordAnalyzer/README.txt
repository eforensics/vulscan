#============================================#
 Word Analyzer 1.0 (with wxPython) by Vulscan
#============================================#

This python program help you find the vulnerability in MS Word file format(DOC) based on OLE(CFB) format.
It can make you easy to understand the OLE structure with TreeCtrl and see the directory related with other things.
Also Word Struct View shows you more detail information about Word FIB values.

Python code : http://code.google.com/p/vulscan/


#============================================#
              Features & note
#============================================#

>> This release(1.0) contains func
	- OLE Directory Properties listctrl
		- simple Hex View (Right click menu)
		- stream dump (Right click menu)
	- Ole Directory Tree
	- Word FIB Struct View

>> This program made with wxPython. So you have to setup up the wxPython module. Of course, 
   there should be installed the python. 
	- wxPython : http://www.wxpython.org/download.php
	- Python : http://www.python.org/download/
	- recommended : Python 2.7 / wxPython2.8(unicode) / Windows

>> As you know, this Program is just start. It has a lot of problem and bug, and inconvenience GUI.
   But if you give us note or comment even very small thing, that makes it a very useful software as your expected. 


#============================================#
                  Files list
#============================================#

>> WordAnalyzer.py
	- Main python code
>> OLE.py

>> OleStructure.py
	- OLE
>> WordStructure.py
	- FIB 
>> mung.png
	- just..puppy pic :-)