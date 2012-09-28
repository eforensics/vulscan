#!/usr/bin/python

__description__ = 'PDF-Analyzer, use it to analyze a malicious PDF document'
__author__ = 'Hong'
__version__ = '0.1.2'
__date__ = '2012/07/27'

import string
import os, sys
import pdf_anal
import optparse
import jsbeautifier
import vulnerability


def Main():
    oParser = optparse.OptionParser(usage='usage: %prog [options] pdf-file\n' + __description__, version='%prog ' + __version__)
    oParser.add_option('-a', '--all', action='store_true', default=False, help='decode all incoded obj')
    oParser.add_option('-o', '--object', help='decode a specified object and save the object to file')
    oParser.add_option('-j', '--js', action='store_true', default=False, help='find JS object')
    oParser.add_option('-e', '--embedded', action='store_true', default=False, help='find embeddedFile object')
    (options, args) = oParser.parse_args()
    
    print(options,args)
    
    if len(args) == 0 : 
        oParser.print_help()
        print ''
        print '  %s' % __description__
        print '  Have a Fun!!!'
        
    else :
        pdf = pdf_anal.pdf_read()        
        f = open(args[0],'rb')        
        pdf.file_read(f)
        
        anal = pdf_anal.find_obj()
        anal.__ini__(pdf.obj_data)
        
        f.seek(0)        
        file_whole_data = f.read()
        f.close()
                
        ff = open('%s_result.txt'%args[0],'wb')
        #fdata = open('%s_anwa.txt'%args[0],'wb')
        
        ff.write("\r\n##### Object No. & Attribution & Stream point in each Object #####\r\n")
        for i in range(len(pdf.obj_data)):
            ff.write(str(pdf.obj_data[i][0])+"\t"+str(pdf.obj_data[i][1])+"\t"+str(pdf.obj_data[i][2])+"\r\n")
            
        if options.js or options.embedded:
            num = []
            print("[*] Trailer Object Searching...")
            if not pdf.trailer_att == "": ## checking the trailer obj
                print("[-] Trailer Object is Found!\r\n")
                print(pdf.trailer_att)
                
                if options.js :
                    if anal.find_main("Trailer", pdf.trailer_att, pdf.obj_tree,'/JS'):
                        num = anal.obj_tree[len(anal.obj_tree)-1][2]## The last list in obj_tree points a obj contained /JS att. That in't stream obj
                else :
                    if anal.find_main("Trailer", pdf.trailer_att, pdf.obj_tree,'/EmbeddedFile'):
                        num = anal.obj_tree[len(anal.obj_tree)-1][0]
                    
                if not num == [] :
                    obj_num = anal.find_obj_data_num(num[0])
                    print("obj Num : %s"%num)
                    #filter = anal.find_filter(pdf.obj_data[obj_num][1])
                    #decoded_data = anal.decoding_string(filter,pdf.obj_data[obj_num][2])
                    print(pdf.obj_data[obj_num][1])
                    decoded_data = anal.decoding_string(pdf.obj_data[obj_num][1],pdf.obj_data[obj_num][2],file_whole_data)
                    
                    if pdf_anal.isJavascript(decoded_data) :
                        print("[*] True return at isJavascript func\n")
                        decoded_data = jsbeautifier.beautify(decoded_data)
                        vulnerability.vulnerability(decoded_data)
                    else :
                        print("[*] False return at isJavascript func\n")
                                                
                    
                    ff.write("\r\n##### Following Objects #####\r\n")
                    for i in range(len(anal.obj_tree)):
                        ff.write(str(anal.obj_tree[i]))
                        ff.write("\r\n")
                    ff.write("\r\n\r\n##### Decoded Data  / Filter : %s / Obj NO. %s #####\r\n"%(str(filter), str(num)))
                    ff.write(decoded_data+"\r\n")
                    
                    fDe_data = open('%s_decoded_data.txt'%args[0],'wb')        
        
                    fDe_data.write('\r\n##### File : %s #####\r\n\r\n'%sys.argv[1])
                    fDe_data.write(decoded_data)
                    
                    fDe_data.close()
                    
                else :
                    print("Nothing...T.T\r\n\r\n")
                
            else :
                print("[-] Trailer Object isn't Found!\r\n")
                
        elif options.all :
            ff.write("\r\n##### Object No. & Attribution & Stream in each Object #####\r\n")
            for i in range(len(pdf.obj_data)):
                #stream = anal.decode_obj(i, file_whole_data)
                stream = anal.decode_obj(i, file_whole_data)
                if stream == 0 :
                    ff.write(str(pdf.obj_data[i][0])+"\t"+str(pdf.obj_data[i][1])+"\r\n"+str(pdf.obj_data[i][2])+"\r\n")
                else :
                    ff.write(str(pdf.obj_data[i][0])+"\t"+str(pdf.obj_data[i][1])+"\t"+str(pdf.obj_data[i][2])+"\r\n"+stream+"\r\n")
                ff.write("\r\n########################################################\r\n")
                
        elif options.object :
            obj_num = anal.find_obj_data_num(options.object)
            if obj_num == -1 :
                print("This Object doesn't exist in file. check again the number\r\n")
            else :
                print(options.object, obj_num)
                de_data = anal.decode_obj(obj_num, file_whole_data)
                
                if de_data == 0:
                    print("decoding error\n")
                else :
                    fObject = open('%s_%s.txt'%(args[0],options.object),'wb')
                    fObject.write(str(de_data))
                    fObject.close()
                
                
        ff.close()
        #fdata.close()

if __name__ == '__main__':
    Main()
