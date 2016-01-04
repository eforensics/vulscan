### **문서 파일내에 취약점이 존재하는지 검사하는 스캐너입니다.** ###
```
■ 대상 파일 : doc, ppt, xls, hwp, pdf ( Version 1.0 )

■ 지원 예정 파일 : xml, rtf

■ 인터페이스 : CLI 방식
```

### **개발 환경** ###
```
■ 파이썬 버전 : 2.7.1

■ 개발툴 : Eclipse

■ 운영체제 : Windows ( Only )
```

### **개발 일정** ###
```
■ Version 0.1 : Office 제품군과 한글 워드 프로세서 문서 파일 Parse / Dump 기능 제공
   - 위치 : /branches/
   - Commit 예정 : 2012.09.18

■ Version 0.2 : Adobe 제품군에 대한 문서 파일 Parse / Dump 기능 제공
   - 위치 : /branches/
   - Commit 예정 : 2012.09.25

■ Version 0.3 : Version 0.1과 Version 0.2의 기능 통합 모듈 제공
   - 위치 : /trunk/
   - Commit 예정 : 2012.10.08

■ Version 1.0 : 취약점 Scanning 기능 제공
   - 위치 : /trunk/
   - Commit 예정 : 2012.11.05


※ x.x 하위의 x.x.x 버전은 해당 버전이 갖는 버그 수정을 의미합니다. 
```

### **참고 문서** ###
```
※ 참고 문서는 해당 Project 페이지의 Downloads 부분을 참고하시기 바랍니다. ( http://code.google.com/p/vulscan/downloads/list )
```

### **참고 모듈** ###
```
[ PDF ]
■ Didier Stevens : pdf-parser.py ( http://blog.didierstevens.com/programs/pdf-tools/ )
     - IsJavaScript( ) 

[ OLE ]
■ Python : OleFileIO_PL ( http://pypi.python.org/pypi/OleFileIO_PL )
```