# -*- coding:utf-8 -*-
#
#
#
#   명명 규칙
#        - Python에서 제공하고 있는 기본 모듈에 대해서는 하위 명명 규칙을 따르지 않는다. 
#        - 변수명 선언
#            a. 변수명 선언은 다음의 형태를 따른다.
#                <Prefix>_<Variable>
#            b. Prefix는 다음의 명명법을 따른다.
#                1) Prefix는 소문자로 구성된다.
#                2) Prefix는 다음의 요소들을 사용한다.
#                    ================================================
#                    Type            Prefix            Example
#                    ------------------------------------------------
#                    Tuple           t                t_OLEHeader
#                    List            l                l_MSAT
#                    Double-List     dl               dl_OLEDirectory
#                    DDouble-List    ddl              ddl_Object
#                    Dictionary      d                d_Format
#                    String          s                s_Buffer
#                    Binary          b                b_Buffer
#                    Number          n                n_Length
#                    Pointer         p                p_file
#                    Bool            B                B_Ret
#
#                    Member          m                m_Member
#                    Class           C                CFile
#                    Class Instance  None             File 
#                    
#                    Function        fn               fnReadFile
#    
#                    Global Data     g                g_Formats
#                    ================================================
#            c. Variable은 다음의 명명법을 따른다. 
#                1) 첫 한글자는 대문자로 시작한다. 
#
#
#
