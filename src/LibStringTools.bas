Attribute VB_Name = "LibStringTools"
'===============================================================================
' VBA StringTools
' ------------------------------------------
' https://github.com/guwidoe/VBA-StringTools
' ------------------------------------------
' MIT License
'
' Copyright (c) 2023 Guido Witt-Döring
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'===============================================================================

'TODO:
'Make ReplaceUnicodeLiterals Mac compatible by removing Regex

Option Explicit
Option Base 0

#Const TEST_MODE = True

#If Mac Then
    #If VBA7 Then 'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
        Private Declare PtrSafe Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As LongPtr, ByVal fromCode As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr) As Long
        Private Declare PtrSafe Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr, ByRef inBuf As LongPtr, ByRef inBytesLeft As LongPtr, ByRef outBuf As LongPtr, ByRef outBytesLeft As LongPtr) As LongPtr
        
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, source As Any, ByVal length As LongPtr) As LongPtr
        Private Declare PtrSafe Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As LongPtr
    #Else
        Private Declare Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long, ByRef inBuf As Long, ByRef inBytesLeft As Long, ByRef outBuf As Long, ByRef outBytesLeft As Long) As Long
        Private Declare Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As Long, ByVal fromCode As Long) As Long
        Private Declare Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long) As Long
        
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
        Private Declare Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
        
        Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
        Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
    
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal source As LongPtr, ByVal length As Long)
        Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    #Else
        Private Declare Function MultiByteToWideChar Lib "kernel32" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
        Private Declare Function WideCharToMultiByte Lib "kernel32" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    
        Private Declare Function GetLastError Lib "kernel32" () As Long
        Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
    
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
        Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    #End If
#End If

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

Private Const WC_ERR_INVALID_CHARS As Long = &H80&
Private Const MB_ERR_INVALID_CHARS As Long = &H8&

Private Const ERROR_INVALID_PARAMETER      As Long = 87
Private Const ERROR_INSUFFICIENT_BUFFER    As Long = 122
Private Const ERROR_INVALID_FLAGS          As Long = 1004
Private Const ERROR_NO_UNICODE_TRANSLATION As Long = 1113

Private Const MAC_API_ERR_EILSEQ As Long = 92 'Illegal byte sequence
Private Const MAC_API_ERR_EINVAL As Long = 22 'Invalid argument
Private Const MAC_API_ERR_E2BIG  As Long = 7  'Argument list too long

Private Const vbErrInternalError As Long = 51

'https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
Public Enum CodePageIdentifier
    [_first] = -1 '(Is initialized)
  'Enum_Name   Identifier             '.NET Name               Additional information
    cpIBM037 = 37                     'IBM037                  IBM EBCDIC US-Canada
    cpIBM437 = 437                    'IBM437                  OEM United States
    cpIBM500 = 500                    'IBM500                  IBM EBCDIC International
    cpASMO_708 = 708                  'ASMO-708                Arabic (ASMO 708)
    cpASMO_449 = 709                  '                        Arabic (ASMO-449+, BCON V4)
    cpTransparent_Arabic = 710        '                        Arabic - Transparent Arabic
    cpDOS_720 = 720                   'DOS-720                 Arabic (Transparent ASMO); Arabic (DOS)
    cpIbm737 = 737                    'ibm737                  OEM Greek (formerly 437G); Greek (DOS)
    cpIbm775 = 775                    'ibm775                  OEM Baltic; Baltic (DOS)
    cpIbm850 = 850                    'ibm850                  OEM Multilingual Latin 1; Western European (DOS)
    cpIbm852 = 852                    'ibm852                  OEM Latin 2; Central European (DOS)
    cpIBM855 = 855                    'IBM855                  OEM Cyrillic (primarily Russian)
    cpIbm857 = 857                    'ibm857                  OEM Turkish; Turkish (DOS)
    cpIBM00858 = 858                  'IBM00858                OEM Multilingual Latin 1 + Euro symbol
    cpIBM860 = 860                    'IBM860                  OEM Portuguese; Portuguese (DOS)
    cpIbm861 = 861                    'ibm861                  OEM Icelandic; Icelandic (DOS)
    cpDOS_862 = 862                   'DOS-862                 OEM Hebrew; Hebrew (DOS)
    cpIBM863 = 863                    'IBM863                  OEM French Canadian; French Canadian (DOS)
    cpIBM864 = 864                    'IBM864                  OEM Arabic; Arabic (864)
    cpIBM865 = 865                    'IBM865                  OEM Nordic; Nordic (DOS)
    cpCp866 = 866                     'cp866                   OEM Russian; Cyrillic (DOS)
    cpIbm869 = 869                    'ibm869                  OEM Modern Greek; Greek, Modern (DOS)
    cpIBM870 = 870                    'IBM870                  IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
    cpWindows_874 = 874               'windows-874             Thai (Windows)
    cpCp875 = 875                     'cp875                   IBM EBCDIC Greek Modern
    cpShift_jis = 932                 'shift_jis               ANSI/OEM Japanese; Japanese (Shift-JIS)
    cpGb2312 = 936                    'gb2312                  ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
    cpKs_c_5601_1987 = 949            'ks_c_5601-1987          ANSI/OEM Korean (Unified Hangul Code)
    cpBig5 = 950                      'big5                    ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
    cpIBM1026 = 1026                  'IBM1026                 IBM EBCDIC Turkish (Latin 5)
    cpIBM01047 = 1047                 'IBM01047                IBM EBCDIC Latin 1/Open System
    cpIBM01140 = 1140                 'IBM01140                IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)
    cpIBM01141 = 1141                 'IBM01141                IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
    cpIBM01142 = 1142                 'IBM01142                IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)
    cpIBM01143 = 1143                 'IBM01143                IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)
    cpIBM01144 = 1144                 'IBM01144                IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)
    cpIBM01145 = 1145                 'IBM01145                IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)
    cpIBM01146 = 1146                 'IBM01146                IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)
    cpIBM01147 = 1147                 'IBM01147                IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)
    cpIBM01148 = 1148                 'IBM01148                IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)
    cpIBM01149 = 1149                 'IBM01149                IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)
    cpUTF_16 = 1200                   'utf-16                  Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
    cpUnicodeFFFE = 1201              'unicodeFFFE             Unicode UTF-16, big endian byte order; available only to managed applications
    cpWindows_1250 = 1250             'windows-1250            ANSI Central European; Central European (Windows)
    cpWindows_1251 = 1251             'windows-1251            ANSI Cyrillic; Cyrillic (Windows)
    cpWindows_1252 = 1252             'windows-1252            ANSI Latin 1; Western European (Windows)
    cpWindows_1253 = 1253             'windows-1253            ANSI Greek; Greek (Windows)
    cpWindows_1254 = 1254             'windows-1254            ANSI Turkish; Turkish (Windows)
    cpWindows_1255 = 1255             'windows-1255            ANSI Hebrew; Hebrew (Windows)
    cpWindows_1256 = 1256             'windows-1256            ANSI Arabic; Arabic (Windows)
    cpWindows_1257 = 1257             'windows-1257            ANSI Baltic; Baltic (Windows)
    cpWindows_1258 = 1258             'windows-1258            ANSI/OEM Vietnamese; Vietnamese (Windows)
    cpJohab = 1361                    'Johab                   Korean (Johab)
    cpMacintosh = 10000               'macintosh               MAC Roman; Western European (Mac)
    cpX_mac_japanese = 10001          'x-mac-japanese          Japanese (Mac)
    cpX_mac_chinesetrad = 10002       'x-mac-chinesetrad       MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
    cpX_mac_korean = 10003            'x-mac-korean            Korean (Mac)
    cpX_mac_arabic = 10004            'x-mac-arabic            Arabic (Mac)
    cpX_mac_hebrew = 10005            'x-mac-hebrew            Hebrew (Mac)
    cpX_mac_greek = 10006             'x-mac-greek             Greek (Mac)
    cpX_mac_cyrillic = 10007          'x-mac-cyrillic          Cyrillic (Mac)
    cpX_mac_chinesesimp = 10008       'x-mac-chinesesimp       MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
    cpX_mac_romanian = 10010          'x-mac-romanian          Romanian (Mac)
    cpX_mac_ukrainian = 10017         'x-mac-ukrainian         Ukrainian (Mac)
    cpX_mac_thai = 10021              'x-mac-thai              Thai (Mac)
    cpX_mac_ce = 10029                'x-mac-ce                MAC Latin 2; Central European (Mac)
    cpX_mac_icelandic = 10079         'x-mac-icelandic         Icelandic (Mac)
    cpX_mac_turkish = 10081           'x-mac-turkish           Turkish (Mac)
    cpX_mac_croatian = 10082          'x-mac-croatian          Croatian (Mac)
    cpUTF_32 = 12000                  'utf-32                  Unicode UTF-32, little endian byte order; available only to managed applications
    cpUTF_32BE = 12001                'utf-32BE                Unicode UTF-32, big endian byte order; available only to managed applications
    cpX_Chinese_CNS = 20000           'x-Chinese_CNS           CNS Taiwan; Chinese Traditional (CNS)
    cpX_cp20001 = 20001               'x-cp20001               TCA Taiwan
    cpX_Chinese_Eten = 20002          'x_Chinese-Eten          Eten Taiwan; Chinese Traditional (Eten)
    cpX_cp20003 = 20003               'x-cp20003               IBM5550 Taiwan
    cpX_cp20004 = 20004               'x-cp20004               TeleText Taiwan
    cpX_cp20005 = 20005               'x-cp20005               Wang Taiwan
    cpX_IA5 = 20105                   'x-IA5                   IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)
    cpX_IA5_German = 20106            'x-IA5-German            IA5 German (7-bit)
    cpX_IA5_Swedish = 20107           'x-IA5-Swedish           IA5 Swedish (7-bit)
    cpX_IA5_Norwegian = 20108         'x-IA5-Norwegian         IA5 Norwegian (7-bit)
    cpUs_ascii = 20127                'us-ascii                US-ASCII (7-bit)
    cpX_cp20261 = 20261               'x-cp20261               T.61
    cpX_cp20269 = 20269               'x-cp20269               ISO 6937 Non-Spacing Accent
    cpIBM273 = 20273                  'IBM273                  IBM EBCDIC Germany
    cpIBM277 = 20277                  'IBM277                  IBM EBCDIC Denmark-Norway
    cpIBM278 = 20278                  'IBM278                  IBM EBCDIC Finland-Sweden
    cpIBM280 = 20280                  'IBM280                  IBM EBCDIC Italy
    cpIBM284 = 20284                  'IBM284                  IBM EBCDIC Latin America-Spain
    cpIBM285 = 20285                  'IBM285                  IBM EBCDIC United Kingdom
    cpIBM290 = 20290                  'IBM290                  IBM EBCDIC Japanese Katakana Extended
    cpIBM297 = 20297                  'IBM297                  IBM EBCDIC France
    cpIBM420 = 20420                  'IBM420                  IBM EBCDIC Arabic
    cpIBM423 = 20423                  'IBM423                  IBM EBCDIC Greek
    cpIBM424 = 20424                  'IBM424                  IBM EBCDIC Hebrew
    cpX_EBCDIC_KoreanExtended = 20833 'x-EBCDIC-KoreanExtended IBM EBCDIC Korean Extended
    cpIBM_Thai = 20838                'IBM-Thai                IBM EBCDIC Thai
    cpKoi8_r = 20866                  'koi8-r                  Russian (KOI8-R); Cyrillic (KOI8-R)
    cpIBM871 = 20871                  'IBM871                  IBM EBCDIC Icelandic
    cpIBM880 = 20880                  'IBM880                  IBM EBCDIC Cyrillic Russian
    cpIBM905 = 20905                  'IBM905                  IBM EBCDIC Turkish
    cpIBM00924 = 20924                'IBM00924                IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)
    cpEuc_jp = 20932                  'EUC-JP                  Japanese (JIS 0208-1990 and 0212-1990)
    cpX_cp20936 = 20936               'x-cp20936               Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)
    cpX_cp20949 = 20949               'x-cp20949               Korean Wansung
    cpCp1025 = 21025                  'cp1025                  IBM EBCDIC Cyrillic Serbian-Bulgarian
    cpDeprecated = 21027                       '                        (deprecated)
    cpKoi8_u = 21866                  'koi8-u                  Ukrainian (KOI8-U); Cyrillic (KOI8-U)
    cpIso_8859_1 = 28591              'iso-8859-1              ISO 8859-1 Latin 1; Western European (ISO)
    cpIso_8859_2 = 28592              'iso-8859-2              ISO 8859-2 Central European; Central European (ISO)
    cpIso_8859_3 = 28593              'iso-8859-3              ISO 8859-3 Latin 3
    cpIso_8859_4 = 28594              'iso-8859-4              ISO 8859-4 Baltic
    cpIso_8859_5 = 28595              'iso-8859-5              ISO 8859-5 Cyrillic
    cpIso_8859_6 = 28596              'iso-8859-6              ISO 8859-6 Arabic
    cpIso_8859_7 = 28597              'iso-8859-7              ISO 8859-7 Greek
    cpIso_8859_8 = 28598              'iso-8859-8              ISO 8859-8 Hebrew; Hebrew (ISO-Visual)
    cpIso_8859_9 = 28599              'iso-8859-9              ISO 8859-9 Turkish
    cpIso_8859_13 = 28603             'iso-8859-13             ISO 8859-13 Estonian
    cpIso_8859_15 = 28605             'iso-8859-15             ISO 8859-15 Latin 9
    cpX_Europa = 29001                'x-Europa                Europa 3
    cpIso_8859_8_i = 38598            'iso-8859-8-i            ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
    cpIso_2022_jp = 50220             'iso-2022-jp             ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
    cpCsISO2022JP = 50221             'csISO2022JP             ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
    cpIso_2022_jp_w_1b_Kana = 50222   'iso-2022-jp             ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
    cpIso_2022_kr = 50225             'iso-2022-kr             ISO 2022 Korean
    cpX_cp50227 = 50227               'x-cp50227               ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
    cpISO_2022_Trad_Chinese = 50229   '                        ISO 2022 Traditional Chinese
    cpEBCDIC_Jap_Katakana_Ext = 50930 '                        EBCDIC Japanese (Katakana) Extended
    cpEBCDIC_US_Can_and_Jap = 50931   '                        EBCDIC US-Canada and Japanese
    cpEBCDIC_Kor_Ext_and_Kor = 50933  '                        EBCDIC Korean Extended and Korean
    cpEBCDIC_Simp_Chin_Ext = 50935    '                        EBCDIC Simplified Chinese Extended and Simplified Chinese
    cpEBCDIC_Simp_Chin = 50936        '                        EBCDIC Simplified Chinese
    cpEBCDIC_US_Can_Trad_Chin = 50937 '                        EBCDIC US-Canada and Traditional Chinese
    cpEBCDIC_Jap_Latin_Ext = 50939    '                        EBCDIC Japanese (Latin) Extended and Japaneseeuc_jp = 51932                  'euc-jp                 EUC Japanese
    cpEUC_CN = 51936                  'EUC-CN                  EUC Simplified Chinese; Chinese Simplified (EUC)
    cpEuc_kr = 51949                  'euc-kr                  EUC Korean
    cpEUC_Traditional_Chinese = 51950 '                        EUC Traditional Chinese
    cpHz_gb_2312 = 52936              'hz-gb-2312              HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
    cpGB18030 = 54936                 'GB18030                 Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
    cpX_iscii_de = 57002              'x-iscii-de              ISCII Devanagari
    cpX_iscii_be = 57003              'x-iscii-be              ISCII Bangla
    cpX_iscii_ta = 57004              'x-iscii-ta              ISCII Tamil
    cpX_iscii_te = 57005              'x-iscii-te              ISCII Telugu
    cpX_iscii_as = 57006              'x-iscii-as              ISCII Assamese
    cpX_iscii_or = 57007              'x-iscii-or              ISCII Odia
    cpX_iscii_ka = 57008              'x-iscii-ka              ISCII Kannada
    cpX_iscii_ma = 57009              'x-iscii-ma              ISCII Malayalam
    cpX_iscii_gu = 57010              'x-iscii-gu              ISCII Gujarati
    cpX_iscii_pa = 57011              'x-iscii-pa              ISCII Punjabi
    cpUTF_7 = 65000                   'utf-7                   Unicode (UTF-7)
    cpUTF_8 = 65001                   'utf-8                   Unicode (UTF-8)
    [_last]
End Enum



'According to documentation:
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
'Note: The documentation doesn't seem to list all codepages for which certain
'      flags are disallowed. This can lead to 'Library implementation erroneous'
'      errors when calling Encode, Decode or Transcode with 'raiseErrors = True'
Private Static Function CodePageAllowsFlags(ByVal cpID As Long) As Boolean
    Dim arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As Boolean
    
    If arr(CodePageIdentifier.[_first]) Then
        CodePageAllowsFlags = arr(cpID)
        Exit Function
    End If
    
    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        arr(i) = True
    Next i
    
    'According to docs:
    arr(cpIso_2022_jp) = False
    arr(cpCsISO2022JP) = False
    arr(cpIso_2022_jp_w_1b_Kana) = False
    arr(cpIso_2022_kr) = False
    arr(cpX_cp50227) = False
    arr(cpISO_2022_Trad_Chinese) = False
    For i = cpX_iscii_de To cpX_iscii_pa
        arr(i) = False
    Next i
    arr(cpUTF_7) = False

    'According to trial and error, it is easier to whitelist:
    For i = CodePageIdentifier.[_first] + 1 To CodePageIdentifier.[_last]
        arr(i) = False
    Next i
    arr(cpUTF_32) = True   'Not sure about this one
    arr(cpUTF_32BE) = True 'Not sure about this one
    arr(cpGB18030) = True  'This one is definitely allowed
    arr(cpUTF_8) = True    'This one is definitely allowed
    
    CodePageAllowsFlags = arr(cpID)
End Function

'According to documentation:
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
Private Static Function CodePageAllowsQueryReversible(ByVal cpID As Long) As Boolean
    Dim arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As Boolean
    
    If arr(CodePageIdentifier.[_first]) Then
        CodePageAllowsQueryReversible = arr(cpID)
        Exit Function
    End If
    
    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        arr(i) = True
    Next i
    
    'According to docs:
    arr(cpUTF_7) = False
    arr(cpUTF_8) = False
    
    'According to trial and error there are a bunch more:
    arr(cpIso_2022_jp) = False
    arr(cpCsISO2022JP) = False
    arr(cpIso_2022_jp_w_1b_Kana) = False
    arr(cpIso_2022_kr) = False
    arr(cpX_cp50227) = False
    arr(cpISO_2022_Trad_Chinese) = False
    arr(cpHz_gb_2312) = False
    arr(cpGB18030) = False
    arr(cpX_iscii_de) = False
    arr(cpX_iscii_be) = False
    arr(cpX_iscii_ta) = False
    arr(cpX_iscii_te) = False
    arr(cpX_iscii_as) = False
    arr(cpX_iscii_or) = False
    arr(cpX_iscii_ka) = False
    arr(cpX_iscii_ma) = False
    arr(cpX_iscii_gu) = False
    arr(cpX_iscii_pa) = False
    
    CodePageAllowsQueryReversible = arr(cpID)
End Function

'Returns an array for converting CodePageIDs to ConversionDescriptorNames
Private Static Function ConvDescriptorName(ByVal cpID As Long) As String
    Dim arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As String
    
    If arr(CodePageIdentifier.[_first]) Then
        ConvDescriptorName = StrConv(arr(cpID), vbFromUnicode)
        Exit Function
    End If
    
    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        arr(i) = -1
    Next i
    
    'Source:
    'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv_open.3.html#//apple_ref/doc/man/3/iconv_open
    'European languages
    arr(cpIso_8859_1) = "ISO-8859-1"
    arr(cpIso_8859_2) = "ISO-8859-2"
    arr(cpIso_8859_3) = "ISO-8859-3"
    arr(cpIso_8859_4) = "ISO-8859-4"
    arr(cpIso_8859_5) = "ISO-8859-5"
    arr(cpIso_8859_7) = "ISO-8859-7"
    arr(cpIso_8859_9) = "ISO-8859-9"
    'arr( ) =  "ISO-8859-10"
    arr(cpIso_8859_13) = "ISO-8859-13"
    'arr( ) =  "ISO-8859-14"
    arr(cpIso_8859_15) = "ISO-8859-15"
    'arr( ) =  "ISO-8859-16"
    'arr( ) =  "KOI8-R"
    arr(cpKoi8_u) = "KOI8-U"
    'arr( ) =  "KOI8-RU"
    arr(cpWindows_1250) = "CP1250"
    arr(cpWindows_1251) = "CP1251"
    arr(cpWindows_1252) = "CP1252"
    arr(cpWindows_1253) = "CP1253"
    arr(cpWindows_1254) = "CP1254"
    arr(cpWindows_1257) = "CP1257"
    arr(cpIbm850) = "CP850"
    arr(cpCp866) = "CP866"
    'arr( ) =  "MacRoman"
    arr(cpX_mac_ce) = "MacCentralEurope"
    arr(cpX_mac_icelandic) = "MacIceland"
    arr(cpX_mac_croatian) = "MacCroatian"
    arr(cpX_mac_romanian) = "MacRomania"
    arr(cpX_mac_cyrillic) = "MacCyrillic"
    arr(cpX_mac_ukrainian) = "MacUkraine"
    arr(cpX_mac_greek) = "MacGreek"
    arr(cpX_mac_turkish) = "MacTurkish"
    arr(cpMacintosh) = "Macintosh"

    'Semitic languages
    arr(cpIso_8859_6) = "ISO-8859-6"
    arr(cpIso_8859_8) = "ISO-8859-8"
    arr(cpWindows_1255) = "CP1255"
    arr(cpWindows_1256) = "CP1256"
    arr(cpDOS_862) = "CP862"
    arr(cpX_mac_hebrew) = "MacHebrew"
    arr(cpX_mac_arabic) = "MacArabic"

    'Japanese
    'arr( ) =  "EUC-JP"
    arr(cpShift_jis) = "SHIFT_JIS"
    arr(cpShift_jis) = "CP932" '(duplicate)
    'arr( ) =  "ISO-2022-JP"
    'arr( ) =  "ISO-2022-JP-2"
    'arr( ) =  "ISO-2022-JP-1"

    'Chinese
    'arr( ) =  "EUC-CN"
    'arr( ) =  "HZ"
    'arr( ) =  "GBK"
    'arr( ) =  "CP936"
    'arr( ) =  "GB18030"
    'arr( ) =  "EUC-TW"
    arr(cpBig5) = "BIG5"
    arr(cpBig5) = "CP950" '(duplicate)
    'arr( ) =  "BIG5-HKSCS"
    'arr( ) =  "BIG5-HKSCS:2001"
    'arr( ) =  "BIG5-HKSCS:1999"
    'arr( ) =  "ISO-2022-CN"
    'arr( ) =  "ISO-2022-CN-EXT"

    'Korean
    arr(cpEuc_kr) = "EUC-KR"
    arr(cpKs_c_5601_1987) = "CP949"
    arr(cpIso_2022_kr) = "ISO-2022-KR"
    'arr( ) =  "JOHAB"

    'Armenian
    'arr( ) =  "ARMSCII-8"

    'Georgian
    'arr( ) =  "Georgian-Academy"
    'arr( ) =  "Georgian-PS"

    'Tajik
    'arr( ) =  "KOI8-T"

    'Kazakh
    'arr( ) =  "PT154"

    'Thai
    'arr( ) =  "TIS-620"
    arr(cpWindows_874) = "CP874"
    arr(cpX_mac_thai) = "MacThai"

    'Laotian
    'arr( ) =  "MuleLao-1"
    'arr( ) =  "CP1133"

    'Vietnamese
    'arr( ) =  "VISCII"
    'arr( ) =  "TCVN"
    arr(cpWindows_1258) = "CP1258"

    'Platform specifics
    'arr( ) =  "HP-ROMAN8"
    'arr( ) =  "NEXTSTEP"

    'Full Unicode
    'arr( ) =  "UCS-2"
    arr(cpUnicodeFFFE) = "UCS-2BE" '(duplicate)
    arr(cpUTF_16) = "UCS-2LE" '(duplicate)
    'arr( ) =  "UCS-4"
    arr(cpUTF_32BE) = "UCS-4BE" '(duplicate)
    arr(cpUTF_32) = "UCS-4LE" '(duplicate)
    'arr( ) =  "UTF-16"
    arr(cpUnicodeFFFE) = "UTF-16BE"
    arr(cpUTF_16) = "UTF-16LE"
    'arr( ) =  "UTF-32"
    arr(cpUTF_32BE) = "UTF-32BE"
    arr(cpUTF_32) = "UTF-32LE"
    arr(cpUTF_7) = "UTF-7"
    arr(cpUTF_8) = "UTF-8"
    'arr( ) =  "C99"
    'arr( ) =  "JAVA"

    'Full Unicode in terms of uint16_t or uint32_t
    '(with machine dependent endianness and alignment)
    'arr( ) =  "UCS-2-INTERNAL"
    'arr( ) =  "UCS-4-INTERNAL"

    'Locale dependent in terms of char or wchar_t
    '(with  machine  dependent  endianness  and  alignment and with
    'semantics depending on the OS and the  current  LC_CTYPE  locale facet)
    'arr( ) =  "char"
    'arr( ) =  "wchar_t"

    'When  configured with the option --enable-extra-encodings
    'it also pro-vides provides vides support for a few extra encodings:

    'European languages
    arr(cpIBM437) = "CP437"
    arr(cpIbm737) = "CP737"
    arr(cpIbm775) = "CP775"
    arr(cpIbm852) = "CP852"
    'arr( ) =  "CP853"
    arr(cpIBM855) = "CP855"
    arr(cpIbm857) = "CP857"
    arr(cpIBM00858) = "CP858"
    arr(cpIBM860) = "CP860"
    arr(cpIbm861) = "CP861"
    arr(cpIBM863) = "CP863"
    arr(cpIBM865) = "CP865"
    arr(cpIbm869) = "CP869"
    'arr( ) =  "CP1125"

    'Semitic languages
    arr(cpIBM864) = "CP864"

    'Japanese
    'arr( ) =  "EUC-JISX0213"
    'arr( ) =  "Shift_JISX0213"
    'arr( ) =  "ISO-2022-JP-3"

    'Chinese
    'arr( ) = "BIG5-2003" '(experimental)

    'Turkmen
    'arr( ) =  "TDS565"

    'Platform specifics
    'arr( ) =  "ATARIST"
    'arr( ) =  "RISCOS-LATIN1"

    'The empty encoding name is equivalent to "char":
    'it denotes the locale dependent character encoding.
    ConvDescriptorName = ConvDescriptorName(cpID)
End Function

''Returns a Collection for converting ConversionDescriptorNames to CodePageIDs
'Private Function ConvDescriptorNameToCodePage() As Collection
'    Static c As Collection
'
'    If Not c Is Nothing Then
'        Set ConvDescriptorNameToCodePage = c
'        Exit Function
'    End If
'
'    Dim cpID As Long
'    Dim conversionDescriptor As Variant
'    Set c = New Collection
'
'    On Error Resume Next
'    For cpID = CodePageIdentifier.[_first_] To CodePageIdentifier.[_last_]
'        conversionDescriptor = CodePageToConvDescriptorName(CStr(cpID))
'        If conversionDescriptor <> "" Then
'            c.Add Item:=cpID, Key:=conversionDescriptor
'        End If
'        conversionDescriptor = ""
'    Next cpID
'
'    cpID = -1
'    For Each conversionDescriptor In CodePageToConvDescriptorName
'        cpID = c(conversionDescriptor)
'        If cpID = -1 Then c.Add Item:=-1, Key:=conversionDescriptor
'        cpID = -1
'    Next conversionDescriptor
'    On Error GoTo 0
'
'    Set ConvDescriptorNameToCodePage = c
'End Function

Private Function GetApiErrorNumber() As Long
    #If Mac Then
        CopyMemory GetApiErrorNumber, ByVal errno_location(), 4
    #Else
        GetApiErrorNumber = Err.LastDllError 'GetLastError
    #End If
End Function

Private Function SetApiErrorNumber(ByVal errNumber As Long) As Long
    #If Mac Then
        CopyMemory ByVal errno_location(), errNumber, 4
    #Else
        SetLastError errNumber
    #End If
End Function

#If Mac = 0 Then
Public Function GetBstrFromWideStringPtr(ByVal lpwString As LongPtr) As String
    Dim length As Long
    If lpwString Then length = lstrlenW(lpwString)
    If length Then
        GetBstrFromWideStringPtr = Space$(length)
        CopyMemory StrPtr(GetBstrFromWideStringPtr), lpwString, length * 2
    End If
End Function
#End If

'This function attempts to transcode 'str' from codepage 'fromCodePage' to
'codepage 'toCodePage' using the appropriate API functions on the platform.
'Calling it with 'raiseErrors = True' will raise an error if either:
'   - the string 'str' contains byte sequences that do not represent a valid
'     string of codepage 'fromCodePage', or
'   - the string contains codepoints that can not be represented in 'toCodePage'
'     and will lead to the insertion of a "default character".
'E.g.: Transcode("°", cpUTF_16, cpUs_ascii, True) will raise an error, because
'      "°" is not an ASCII character.
'Note that even calling the function with 'raiseErrors = True' doesn't guarantee
'that the conversion is reversible, because sometimes codepoints are replaced
'with more generic characters that aren't the default character (raise no error)
'E.g.:Decode(Transcode("³", cpUTF_16, cpUs_ascii, True), cpUs_ascii) returns "3"
Public Function Transcode(ByRef str As String, _
                          ByVal fromCodePage As CodePageIdentifier, _
                          ByVal toCodePage As CodePageIdentifier, _
                 Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "Transcode"
    'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
    #If Mac Then
        Dim inBytesLeft As LongPtr:  inBytesLeft = LenB(str)
        Dim outBytesLeft As LongPtr: outBytesLeft = inBytesLeft * 4
        Dim buffer As String:        buffer = Space$(CLng(inBytesLeft) * 2)
        Dim inBuf As LongPtr:        inBuf = StrPtr(str)
        Dim outBuf As LongPtr:       outBuf = StrPtr(buffer)
        Dim cd As LongPtr: cd = GetConversionDescriptor(fromCodePage, toCodePage)
        Dim irrevConvCount As Long
        Dim replacementChar As String
        
        Do While inBytesLeft > 0
            SetApiErrorNumber = 0
            irrevConvCount = iconv(cd, inBuf, inBytesLeft, outBuf, outBytesLeft)

            If irrevConvCount = -1 Then 'Error occurred
                If StrPtr(replacementChar) = 0 Then _
                    replacementChar = GetReplacementCharForCodePage(toCodePage)
                
                Select Case GetApiErrorNumber
                    Case MAC_API_ERR_EILSEQ
                        If raiseErrors Then Err.Raise 5, methodName, _
                            "Input is invalid byte sequence of " & _
                            "CodePage " & fromCodePage
                            
                        CopyMemory ByVal outBuf, replacementChar(0), 3
                        outBuf = outBuf + LenB(replacementChar)
                        outBytesLeft = outBytesLeft - LenB(replacementChar)
                        inBuf = inBuf + 1
                        inBytesLeft = inBytesLeft - 1
                    Case MAC_API_ERR_EINVAL
                        If raiseErrors Then Err.Raise 5, methodName, _
                            "Input is incomplete byte sequence of" & _
                            "CodePage " & fromCodePage
                            
                        CopyMemory ByVal outBuf, replacementChar(0), 3
                        outBuf = outBuf + outBytesLeft
                        inBuf = inBuf + inBytesLeft
                        outBytesLeft = 0
                        inBytesLeft = 0
                End Select
            End If
        Loop
        
        If irrevConvCount > 0 And raiseErrors Then Err.Raise 5, _
            methodName, "Default char would be used, encoding would be irreversible"

        Transcode = LeftB$(buffer, LenB(buffer) - CLng(outBytesLeft))

        'These errors are bugs and should be raised even if raiseErrors = False:
        Select Case GetApiErrorNumber
            Case MAC_API_ERR_E2BIG
                Err.Raise vbErrInternalError, methodName, _
                    "Output buffer overrun while transcoding from CodePage " _
                    & fromCodePage & " to CodePage " & toCodePage
            Case Is <> 0
                Err.Raise vbErrInternalError, methodName, "Unknown error " & _
                    "occurred during transcoding with 'iconv'. API Error" & _
                    "Code: " & GetApiErrorNumber
        End Select
        If iconv_close(cd) <> 0 Then
            Err.Raise vbErrInternalError, methodName, "Unknown error occurred" _
                & " when calling 'iconv_close'. API ErrorCode: " & _
                GetApiErrorNumber
        End If
    #Else
        If toCodePage = cpUTF_16 Then
            Transcode = Decode(str, fromCodePage, raiseErrors)
        ElseIf fromCodePage = cpUTF_16 Then
            Transcode = Encode(str, toCodePage, raiseErrors)
        Else
            Transcode = Encode(Decode(str, fromCodePage, raiseErrors), _
                               toCodePage, raiseErrors)
        End If
    #End If
End Function

#If Mac Then
Private Function GetConversionDescriptor( _
                            ByVal fromCodePage As CodePageIdentifier, _
                            ByVal toCodePage As CodePageIdentifier) As LongPtr
    Dim toCpCdName As String:   toCpCdName = ConvDescriptorName(toCodePage)
    Dim fromCpCdName As String: fromCpCdName = ConvDescriptorName(fromCodePage)
    'Todo: potentially implement custom error numbers
    If LenB(toCpCdName) = 0 Then Err.Raise 5, methodName, _
        "No conversion descriptor name assigned to CodePage " & toCodePage
    If LenB(fromCpCdName) = 0 Then Err.Raise 5, methodName, _
        "No conversion descriptor name assigned to CodePage " & fromCodePage
        
    SetApiErrorNumber = 0 'Clear previous errors
    GetConversionDescriptor = iconv_open(StrPtr(toCpCdName), StrPtr(fromCodePage))

    If Not GetConversionDescriptor Then
        Select Case GetApiErrorNumber
            Case MAC_API_ERR_EINVAL
                Err.Raise 5, methodName, "The conversion from CodePage " & _
                    fromCodePage & " to CodePage " & toCodePage & " is not " & _
                    "supported by the implementation of 'iconv' on this platform"
            Case Is <> 0
                Err.Raise vbErrInternalError, methodName, "Unknown error " & _
                    "trying to create a conversion descriptor. API Error" & _
                    "Code: " & GetApiErrorNumber
        End Select
    End If
End Function
#End If

#If Mac Then
'On Mac, replacement character must be manually inserted when using iconv
Private Function GetReplacementCharForCodePage( _
                                    ByVal cpID As CodePageIdentifier) As String
    Static replacementChars As Collection
    If replacementChars Is Nothing Then replacementChars = New Collection
    On Error GoTo 0
    On Error Resume Next
    GetReplacementCharForCodePage = replacementChars(CStr(cpID))
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' U+FFFD (0xEF 0xBF 0xBD) in UTF-8
    Dim ReplacementCharUtf8() As Byte: ReDim ReplacementCharUtf8(0 To 2)
    ReplacementCharUtf8(0) = &HEF
    ReplacementCharUtf8(1) = &HBF
    ReplacementCharUtf8(2) = &HBD
    
    replacementChars.Add Transcode(CStr(ReplacementCharUtf8), cpUTF_8, cpID), _
                         CStr(cpID)
    GetReplacementCharForCodePage = replacementChars(CStr(cpID))
End Function
#End If

'This function tries to encode utf16leStr from vba-internal codepage UTF-16LE to
'codepage 'toCodePage' using the appropriate API functions on the platform.
'Calling it with 'raiseErrors = True' will raise an error if either:
'   - the string 'utf16leStr' contains byte sequences that do not represent a
'     valid UTF-16LE string, or
'   - the string contains codepoints that can not be represented in 'toCodePage'
'     and will lead to the insertion of a "default character".
'E.g.: Encode("°", cpUs_ascii, True) will raise an error, because
'      "°" is not an ASCII character.
'Note that even calling the function with 'raiseErrors = True' doesn't guarantee
'that the conversion is reversible, because sometimes codepoints are replaced
'with more generic characters that aren't the default character (raise no error)
'E.g.: Decode(Encode("³", cpUTF_16, cpUs_ascii, True), cpUs_ascii) returns "3"
Public Function Encode(ByRef utf16leStr As String, _
                       ByVal toCodePage As CodePageIdentifier, _
              Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "Encode"
    
    If toCodePage = cpUTF_16 Then Err.Raise 5, methodName, _
        "Input string should already be UTF-16. Can't encode UTF-16 to UTF-16."
    
    If utf16leStr = vbNullString Then Exit Function
    #If Mac Then
        Encode = Transcode(utf16leStr, cpUTF_16, toCodePage, raiseErrors)
    #Else
        Dim byteCount As Long
        Dim dwFlags As Long
        Dim usedDefaultChar As Boolean
        Dim lpUsedDefaultChar As LongPtr
        If raiseErrors And CodePageAllowsQueryReversible(toCodePage) Then _
            lpUsedDefaultChar = VarPtr(usedDefaultChar)
            
        If raiseErrors And CodePageAllowsFlags(toCodePage) Then _
            dwFlags = WC_ERR_INVALID_CHARS

        SetApiErrorNumber 0
        byteCount = WideCharToMultiByte(toCodePage, dwFlags, StrPtr(utf16leStr), _
                                    Len(utf16leStr), 0, 0, 0, lpUsedDefaultChar)
        If byteCount = 0 Then
            Select Case GetApiErrorNumber
                Case ERROR_NO_UNICODE_TRANSLATION
                    Err.Raise 5, methodName, _
                        "Input is invalid byte sequence of CodePage " & cpUTF_16
                Case ERROR_INVALID_PARAMETER
                    Err.Raise 5, methodName, _
                        "Conversion to CodePage " & toCodePage & " is not " & _
                        "supported by the API on this platform."
                Case ERROR_INSUFFICIENT_BUFFER, ERROR_INVALID_FLAGS
                    Err.Raise vbErrInternalError, methodName, _
                        "Library implementation erroneous. API Error: " & _
                        GetApiErrorNumber
                Case Else
                    Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
            End Select
        End If
        
        If raiseErrors And usedDefaultChar Then _
            Err.Raise 5, methodName, "Default char would be used, encoding " & _
                "would be irreversible."
    
        Dim b() As Byte: ReDim b(0 To byteCount - 1)
        Encode = b
        WideCharToMultiByte toCodePage, dwFlags, StrPtr(utf16leStr), _
                Len(utf16leStr), StrPtr(Encode), byteCount, 0, lpUsedDefaultChar
                
        Select Case GetApiErrorNumber
            Case Is <> 0
                Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
        End Select
    #End If
End Function

'This function tries to decode 'str' from codepage 'fromCodePage' to the vba-
'internal codepage UTF-16LE using the appropriate API functions on the platform.
'Calling it with 'raiseErrors = True' will raise an error if the string 'str'
'contains byte sequences that does not represent a valid encoding in codepage
'fromCodePage.
'E.g.: If 'str' is an UTF-8 encoded string that was read from an external file
'      using 'Open' and 'Get', you can convert it to the VBA-internal UTF-16LE
'      like this:
'      Decode(str, cpUTF_8)
'      If you are afraid 'str' might contain invalid UTF-8 data, use it like so:
'      Decode(str, cpUTF_8, True)
'      The function will now raise an error if invalid UTF-8 data is encountered
Public Function Decode(ByRef str As String, _
                       ByVal fromCodePage As CodePageIdentifier, _
              Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "Decode"
    
    If fromCodePage = cpUTF_16 Then Err.Raise 5, methodName, _
        "VBA strings are UTF-16 by default. No need to decode string from UTF-16."

    If str = vbNullString Then Exit Function
    #If Mac Then
        Decode = Transcode(str, fromCodePage, cpUTF_16, raiseErrors)
    #Else
        Dim charCount As Long
        Dim dwFlags As Long
        
        SetApiErrorNumber 0
        If raiseErrors And CodePageAllowsFlags(fromCodePage) Then _
            dwFlags = MB_ERR_INVALID_CHARS
        
        charCount = MultiByteToWideChar(fromCodePage, dwFlags, StrPtr(str), _
                                        LenB(str), 0, 0)
        If charCount = 0 Then
            Select Case GetApiErrorNumber
                Case ERROR_NO_UNICODE_TRANSLATION
                    Err.Raise 5, methodName, _
                        "Input is invalid byte sequence of CodePage " & cpUTF_16
                Case ERROR_INVALID_PARAMETER
                    Err.Raise 5, methodName, _
                        "Conversion from CodePage " & fromCodePage & " is not" _
                        & " supported by the API on this platform."
                Case ERROR_INSUFFICIENT_BUFFER, ERROR_INVALID_FLAGS
                    Err.Raise vbErrInternalError, methodName, _
                        "Library implementation erroneous. API Error: " & _
                        GetApiErrorNumber
                Case Else
                    Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
            End Select
        End If
    
        Decode = Space$(charCount)
        MultiByteToWideChar fromCodePage, dwFlags, StrPtr(str), LenB(str), _
                            StrPtr(Decode), charCount
                            
        Select Case GetApiErrorNumber
            Case Is <> 0
                Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
        End Select
    #End If
End Function

'Returns strings defined as hex literal as string
'Accepts the following formattings:
'   0xXXXXXX...
'   &HXXXXXX...
'   XXXXXX...
'Where:
'   - prefixes 0x and &H are case sensitive
'   - there's an even number of Xes, X = 0-9 or a-f or A-F (case insensitive)
'Raises error 5 if:
'   - Length is not even / partial bytes
'   - Invalid characters are found (outside prefix and 0-9 / a-f / A-F ranges)
'Examples:
'   - HexToString("0x610062006300") returns "abc"
'   - StrConv(HexToString("0x616263"), vbUnicode) returns "abc"
'   - HexToString("0x61626t") or HexToString("0x61626") both raise error 5
Public Function HexToString(ByRef hexStr As String) As String
    Const methodName As String = "HexToString"
    Const errPrefix As String = "Invalid Hex string literal. "
    Dim size As Long: size = Len(hexStr)
    
    If size = 0 Then Exit Function
    If size Mod 2 = 1 Then Err.Raise 5, methodName, errPrefix & "Uneven length"
    
    Static nibbleMap(0 To 255) As Long 'Nibble: 0 to F. Byte: 00 to FF
    Static charMap(0 To 255) As String
    Dim i As Long
    
    If nibbleMap(0) = 0 Then
        For i = 0 To 255
            nibbleMap(i) = -256 'To force invalid character code
            charMap(i) = ChrB$(i)
        Next i
        For i = 0 To 9
            nibbleMap(Asc(CStr(i))) = i
        Next i
        For i = 10 To 15
            nibbleMap(i + 55) = i 'Asc("A") to Asc("F")
            nibbleMap(i + 87) = i 'Asc("a") to Asc("f")
        Next i
    End If
    
    Dim prefix As String: prefix = Left$(hexStr, 2)
    Dim startPos As Long: startPos = -4 * CLng(prefix = "0x" Or prefix = "&H")
    Dim b() As Byte:      b = hexStr
    Dim j As Long
    Dim charCode As Long
    
    HexToString = MidB$(hexStr, 1, size / 2 - Sgn(startPos))
    For i = startPos To UBound(b) Step 4
        j = j + 1
        charCode = nibbleMap(b(i)) * &H10& + nibbleMap(b(i + 2))
        If charCode < 0 Or b(i + 1) > 0 Or b(i + 3) > 0 Then
            Err.Raise 5, methodName, errPrefix & "Expected a-f/A-F or 0-9"
        End If
        MidB$(HexToString, j, 1) = charMap(charCode)
    Next i
End Function

'Converts the input string into a string of hex literals.
'e.g.: "abc" will be turned into "0x610062006300" (UTF-16LE)
'e.g.: StrConv("ABC", vbFromUnicode) will be turned into "0x414243"
Public Function StringToHex(ByRef s As String) As String
    Static map(0 To 255) As String
    Dim b() As Byte: b = s
    Dim i As Long
    
    If LenB(map(0)) = 0 Then
        For i = 0 To 255
            map(i) = Right$("0" & Hex$(i), 2)
        Next i
    End If
    
    StringToHex = Space$(LenB(s) * 2 + 2)
    Mid$(StringToHex, 1, 2) = "0x"
    
    For i = LBound(b) To UBound(b)
        Mid$(StringToHex, (i + 1) * 2 + 1, 2) = map(b(i))
    Next i
End Function

#If Mac = 0 Then
'Replaces all occurences of unicode literals
'Accepts the following formattings:
'   \uXXXX \uXXXXXXXX    (4 or 8 hex digits, 8 for chars outside BMP)
'   \u{XXXXX} \U{XXXXX}  (1 to 5 hex digits)
'   u+XXXX u+XXXXX       (4 or 5 hex digits)
'   &#dddddd;            (1 to 6 decimal digits)
'Where:
'   - prefixes \u is case insensitive
'   - X = 0-9 or a-f or A-F (also case insensitive)
'Example:
'   - "abcd &#97;u+0062\U0063xy\u{64}" returns "abcd abcxyd"
'Notes:
'   - Avoid u+XXXX syntax if string contains literals without delimiters as it
'     can be misinterpreted if adjacent to text starting with 0-9 or a-f.
'   - This function can be slow for very long input strings with many
'     different literals
Public Function ReplaceUnicodeLiterals(ByRef str As String) As String
    Const PATTERN_UNICODE_LITERALS As String = _
        "\\u000[0-9a-f]{5}|\\u[0-9a-f]{4}|" & _
        "\\u{[0-9a-f]{1,5}}|u\+[0-9|a-f]{4,5}|&#\d{1,6};"
    Dim mc As Object
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = PATTERN_UNICODE_LITERALS
        Set mc = .Execute(str)
    End With
    
    Dim match As Variant
    Dim mv As String
    Dim codepoint As Long
    
    For Each match In mc
        mv = match.Value
        If Left$(mv, 1) = "&" Then
            codepoint = CLng(Mid$(mv, 3, Len(mv) - 3))
        Else
            If Mid$(mv, 3, 1) = "{" Then
                codepoint = CLng("&H" & Mid$(mv, 4, Len(mv) - 4))
            Else
                codepoint = CLng("&H" & Mid$(mv, 3))
            End If
        End If
        If codepoint < &H110000 Then
            If codepoint < &HD800& Or codepoint >= &HE000& Then _
                str = Replace(str, mv, ChrU(codepoint))
        End If
    Next match
    ReplaceUnicodeLiterals = str
End Function
#End If

'Replaces all occurences of unicode characters outside the ANSI codePoint range
'(which the VBA-IDE can not display) with literals of the following formattings:
'   \uXXXX      for characters inside the basic multilingual plane
'   \uXXXXXXXX  for characters outside the basic multilingual plane
'Where:
'   Xes are the digits of the codepoint in hexadecimal. (X = 0-9 or A-F)
Public Function EncodeUnicodeCharacters(ByRef str As String) As String
    Dim codepoint As Long
    Dim i As Long
    Dim j As Long:          j = 1
    Dim result() As String: ReDim result(1 To Len(str))
    
    For i = 1 To Len(str)
        codepoint = AscW(Mid$(str, i, 1)) And &HFFFF&
        
        If codepoint >= &HD800& Then codepoint = AscU(Mid$(str, i, 2))
        
        If codepoint > &HFFFF& Then 'Outside BMP
            result(j) = "\u" & "000" & Hex(codepoint)
            i = i + 1
        ElseIf codepoint > &HFF Then 'BMP
            result(j) = "\u" & Right$("00" & Hex(codepoint), 4)
        Else
            result(j) = Mid$(str, i, 1)
        End If
        j = j + 1
    Next i
    EncodeUnicodeCharacters = Join(result, "")
End Function

'Returns the given unicode codepoint as standard VBA UTF-16LE string
 Public Function ChrU(ByVal codepoint As Long, _
             Optional ByVal allowSingleSurrogates As Boolean = False) _
                      As String
    Const methodName As String = "ChrU"
    
    If codepoint < &H8000 Then Err.Raise 5, methodName, _
        "Argument 'codepoint' = " & codepoint & " < -32768, invalid"
    
    If codepoint < 0 Then codepoint = codepoint And &HFFFF& 'Incase of uInt input
    
    If codepoint < &HD800& Then
        ChrU = ChrW$(codepoint)
    ElseIf codepoint < &HE000& And Not allowSingleSurrogates Then
        Err.Raise 5, methodName, _
            "Invalid Unicode codepoint. (Range reserved for surrogate pairs)"
    ElseIf codepoint < &H10000 Then
        ChrU = ChrW$(codepoint)
    ElseIf codepoint < &H110000 Then
        codepoint = codepoint - &H10000
        ChrU = ChrW$(&HD800& Or (codepoint \ &H400&)) & _
               ChrW$(&HDC00& Or (codepoint And &H3FF&))
    Else
        Err.Raise 5, methodName, "Codepoint outside of valid Unicode range."
    End If
End Function

'Returns a given characters unicode codepoint as long.
'Note: One unicode character can consist of two VBA "characters", a so-called
'      "surrogate pair" (input string of length 2, so Len(char) = 2!)
Public Function AscU(ByRef char As String) As Long
    Dim s As String
    Dim lo As Long
    Dim hi As Long
    
    If Len(char) = 1 Then
        AscU = AscW(char) And &HFFFF&
    Else
        s = Left$(char, 2)
        hi = AscW(Mid$(s, 1, 1)) And &HFFFF&
        lo = AscW(Mid$(s, 2, 1)) And &HFFFF&
        
        If &HDC00& > lo Or lo > &HDFFF& Then
            AscU = hi
            Exit Function
        End If
        AscU = (hi - &HD800&) * &H400& + (lo - &HDC00&) + &H10000
    End If
End Function

'Function transcoding an ANSI encoded string to the VBA-native UTF-16LE
Public Function DecodeANSI(ByRef ansiStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim ansi() As Byte:    ansi = ansiStr
    Dim utf16le() As Byte: ReDim utf16le(0 To LenB(ansiStr) * 2 - 1)
    
    For i = LBound(ansi) To UBound(ansi)
        utf16le(j) = ansi(i)
        j = j + 2
    Next i
    DecodeANSI = utf16le
End Function

'Function transcoding a VBA-native UTF-16LE encoded string to an ANSI string
'Note: Information will be lost for codepoints > 255!
Public Function EncodeANSI(ByRef utf16leStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim utf16le() As Byte: utf16le = utf16leStr
    Dim ansi() As Byte

    ReDim ansi(1 To Len(utf16leStr))
    For i = LBound(ansi) To UBound(ansi)
        If utf16le(j + 1) = 0 Then
            ansi(i) = utf16le(j)
            j = j + 2
        Else
            ansi(i) = &H3F 'Chr(&H3F) = "?"
            j = j + 2
        End If
    Next i
    EncodeANSI = ansi
End Function

Public Function EncodeUTF8(ByRef utf16leStr As String, _
                  Optional ByVal raiseErrors As Boolean = False) As String
    If Len(utf16leStr) < 50 Then
        EncodeUTF8 = EncodeUTF8native(utf16leStr, raiseErrors)
    Else
        EncodeUTF8 = Encode(utf16leStr, cpUTF_8, raiseErrors)
    End If
End Function

Public Function DecodeUTF8(ByRef utf8Str As String, _
                  Optional ByVal raiseErrors As Boolean = False) As String
    If Len(utf8Str) < 50 Then
        DecodeUTF8 = DecodeUTF8native(utf8Str, raiseErrors)
    Else
        DecodeUTF8 = Decode(utf8Str, cpUTF_8, raiseErrors)
    End If
End Function

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-8
#If TEST_MODE Then
Public Function EncodeUTF8native(ByRef utf16leStr As String, _
                         Optional ByVal raiseErrors As Boolean = False) _
                                  As String
#Else
Private Function EncodeUTF8native(ByRef utf16leStr As String, _
                         Optional ByVal raiseErrors As Boolean = False) _
                                  As String
#End If
    Const methodName As String = "EncodeUTF8native"
    Dim codepoint As Long
    Dim lowSurrogate As Long
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0
    Dim utf8() As Byte:       ReDim utf8(Len(utf16leStr) * 4 - 1)
    
    Do While i <= Len(utf16leStr)
        codepoint = AscW(Mid$(utf16leStr, i, 1)) And &HFFFF&
        
        If codepoint >= &HD800& And codepoint <= &HDBFF& Then 'high surrogate
            lowSurrogate = AscW(Mid$(utf16leStr, i + 1, 1)) And &HFFFF&
            
            If &HDC00& <= lowSurrogate And lowSurrogate <= &HDFFF& Then
                codepoint = (codepoint - &HD800&) * &H400& + _
                            (lowSurrogate - &HDC00&) + &H10000
                i = i + 1
            Else
                If raiseErrors Then _
                    Err.Raise 5, methodName, _
                        "Invalid Unicode codepoint. (Lonely high surrogate)"
                codepoint = &HFFFD&
            End If
        End If
        
        If codepoint < &H80& Then
            utf8(j) = codepoint
            j = j + 1
        ElseIf codepoint < &H800& Then
            utf8(j) = &HC0& Or ((codepoint And &H7C0&) \ &H40&)
            utf8(j + 1) = &H80& Or (codepoint And &H3F&)
            j = j + 2
        ElseIf codepoint < &HDC00 Then
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        ElseIf codepoint < &HE000 Then
            If raiseErrors Then _
                Err.Raise 5, methodName, _
                    "Invalid Unicode codepoint. (Lonely low surrogate)"
            codepoint = &HFFFD&
        ElseIf codepoint < &H10000 Then
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        Else
            utf8(j) = &HF0& Or ((codepoint And &H1C0000) \ &H40000)
            utf8(j + 1) = &H80& Or ((codepoint And &H3F000) \ &H1000&)
            utf8(j + 2) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 3) = &H80& Or (codepoint And &H3F&)
            j = j + 4
        End If
        
        i = i + 1
    Loop
    EncodeUTF8native = MidB$(utf8, 1, j)
End Function

'Function transcoding an UTF-8 encoded string to the VBA-native UTF-16LE
'Function transcoding an VBA-native UTF-16LE encoded string to UTF-8
#If TEST_MODE Then
Public Function DecodeUTF8native(ByRef utf8Str As String, _
                   Optional ByVal raiseErrors As Boolean = False) As String
#Else
Private Function DecodeUTF8native(ByRef utf8Str As String, _
                   Optional ByVal raiseErrors As Boolean = False) As String
#End If

    Const methodName As String = "DecodeUTF8native"
    Dim i As Long
    Dim numBytesOfCodePoint As Byte
    
    Static numBytesOfCodePoints(0 To 255) As Byte
    Static mask(2 To 4) As Long
    Static minCp(2 To 4) As Long
    
    If numBytesOfCodePoints(0) = 0 Then
        For i = &H0& To &H7F&: numBytesOfCodePoints(i) = 1: Next i '0xxxxxxx
        '110xxxxx - C0 and C1 are invalid (overlong encoding)
        For i = &HC2& To &HDF&: numBytesOfCodePoints(i) = 2: Next i
        For i = &HE0& To &HEF&: numBytesOfCodePoints(i) = 3: Next i '1110xxxx
       '11110xxx - 11110100, 11110101+ (= &HF5+) outside of valid Unicode range
        For i = &HF0& To &HF4&: numBytesOfCodePoints(i) = 4: Next i
        For i = 2 To 4: mask(i) = (2 ^ (7 - i) - 1): Next i
        minCp(2) = &H80&: minCp(3) = &H800&: minCp(4) = &H10000
    End If
    
    Dim codepoint As Long
    Dim currByte As Byte
    Dim utf8() As Byte:  utf8 = utf8Str
    Dim utf16() As Byte: ReDim utf16(0 To (UBound(utf8) - LBound(utf8) + 1) * 2)
    Dim j As Long:       j = 0
    Dim k As Long

    i = LBound(utf8)
    Do While i <= UBound(utf8)
        codepoint = utf8(i)
        numBytesOfCodePoint = numBytesOfCodePoints(codepoint)
        
        If numBytesOfCodePoint = 0 Then
            If raiseErrors Then Err.Raise 5, methodName, "Invalid byte"
            GoTo insertErrChar
        ElseIf numBytesOfCodePoint = 1 Then
            utf16(j) = codepoint
            j = j + 2
        ElseIf i + numBytesOfCodePoint - 1 > UBound(utf8) Then
            If raiseErrors Then Err.Raise 5, methodName, _
                    "Incomplete UTF-8 codepoint at end of string."
            GoTo insertErrChar
        Else
            codepoint = utf8(i) And mask(numBytesOfCodePoint)
            
            For k = 1 To numBytesOfCodePoint - 1
                currByte = utf8(i + k)
                
                If (currByte And &HC0&) = &H80& Then
                    codepoint = (codepoint * &H40&) + (currByte And &H3F)
                Else
                    If raiseErrors Then _
                        Err.Raise 5, methodName, "Invalid continuation byte"
                    GoTo insertErrChar
                End If
            Next k
            'Convert the Unicode codepoint to UTF-16LE bytes
            If codepoint < minCp(numBytesOfCodePoint) Then
                If raiseErrors Then Err.Raise 5, methodName, "Overlong encoding"
                GoTo insertErrChar
            ElseIf codepoint < &HD800& Then
                utf16(j) = CByte(codepoint And &HFF&)
                utf16(j + 1) = CByte(codepoint \ &H100&)
                j = j + 2
            ElseIf codepoint < &HE000& Then
                If raiseErrors Then Err.Raise 5, methodName, _
                "Invalid Unicode codepoint.(Range reserved for surrogate pairs)"
                GoTo insertErrChar
            ElseIf codepoint < &H10000 Then
                If codepoint = &HFEFF& Then GoTo nextCp '(BOM - will be ignored)
                utf16(j) = codepoint And &HFF&
                utf16(j + 1) = codepoint \ &H100&
                j = j + 2
            ElseIf codepoint < &H110000 Then 'Calculate surrogate pair
                Dim m As Long:           m = codepoint - &H10000
                Dim loSurrogate As Long: loSurrogate = &HDC00& Or (m And &H3FF)
                Dim hiSurrogate As Long: hiSurrogate = &HD800& Or (m \ &H400&)
                
                utf16(j) = hiSurrogate And &HFF&
                utf16(j + 1) = hiSurrogate \ &H100&
                utf16(j + 2) = loSurrogate And &HFF&
                utf16(j + 3) = loSurrogate \ &H100&
                j = j + 4
            Else
                If raiseErrors Then Err.Raise 5, methodName, _
                        "Codepoint outside of valid Unicode range"
insertErrChar:  utf16(j) = &HFD
                utf16(j + 1) = &HFF
                j = j + 2
                
                If numBytesOfCodePoint = 0 Then numBytesOfCodePoint = 1
            End If
        End If
nextCp: i = i + numBytesOfCodePoint 'Move to the next UTF-8 codepoint
    Loop
    DecodeUTF8native = MidB$(utf16, 1, j)
End Function

#If Mac = 0 Then
'Transcoding a VBA-native UTF-16LE encoded string to UTF-8 using ADODB.Stream
'Much faster than EncodeUTF8native, but only available on Windows
#If TEST_MODE Then
Public Function EncodeUTF8usingAdodbStream(ByRef utf16leStr As String) _
                                            As String
#Else
Private Function EncodeUTF8usingAdodbStream(ByRef utf16leStr As String) _
                                            As String
#End If
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText utf16leStr
        .Position = 0
        .Type = 1 ' adTypeBinary
        .Position = 3 ' Skip BOM (Byte Order Mark)
        EncodeUTF8usingAdodbStream = .Read
        .Close
    End With
End Function

'Transcoding an UTF-8 encoded string to VBA-native UTF-16LE using ADODB.Stream
'Faster than DeocdeUTF8native for some strings but only available on Windows
'Warning: This function performs extremely slow for strings bigger than ~5MB
#If TEST_MODE Then
Public Function DecodeUTF8usingAdodbStream(ByRef utf8Str As String) As String
#Else
Private Function DecodeUTF8usingAdodbStream(ByRef utf8Str As String) As String
#End If
    Dim b() As Byte: b = utf8Str
    With CreateObject("ADODB.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .Write b
        .Position = 0
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        DecodeUTF8usingAdodbStream = .ReadText
        .Close
    End With
End Function
#End If

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-32
Public Function EncodeUTF32LE(ByRef utf16leStr As String, _
                     Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "EncodeUTF32LE"
    
    If utf16leStr = "" Then Exit Function
    
    Dim codepoint As Long
    Dim lowSurrogate As Long
    Dim utf32() As Byte:      ReDim utf32(Len(utf16leStr) * 4 - 1)
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0
    
    Do While i <= Len(utf16leStr)
        codepoint = AscW(Mid$(utf16leStr, i, 1)) And &HFFFF&
        
        If codepoint >= &HD800& And codepoint <= &HDBFF& Then 'high surrogate
            lowSurrogate = AscW(Mid$(utf16leStr, i + 1, 1)) And &HFFFF&
            
            If &HDC00& <= lowSurrogate And lowSurrogate <= &HDFFF& Then
                codepoint = (codepoint - &HD800&) * &H400& + _
                            (lowSurrogate - &HDC00&) + &H10000
                i = i + 1
            Else
                If raiseErrors Then Err.Raise 5, methodName, _
                    "Invalid Unicode codepoint. (Lonely high surrogate)"
                codepoint = &HFFFD&
            End If
        End If
        
        If codepoint >= &HD800& And codepoint < &HE000& Then
            If raiseErrors Then Err.Raise 5, methodName, _
                "Invalid Unicode codepoint. (Lonely low surrogate)"
            codepoint = &HFFFD&
        ElseIf codepoint > &H10FFFF Then
            If raiseErrors Then Err.Raise 5, methodName, _
                "Codepoint outside of valid Unicode range"
            codepoint = &HFFFD&
        End If
        
        utf32(j) = codepoint And &HFF&
        utf32(j + 1) = (codepoint \ &H100&) And &HFF&
        utf32(j + 2) = (codepoint \ &H10000) And &HFF&
        i = i + 1: j = j + 4
    Loop
    EncodeUTF32LE = MidB$(utf32, 1, j)
End Function

'Function transcoding an UTF-32 encoded string to the VBA-native UTF-16LE
Public Function DecodeUTF32LE(ByRef utf32str As String, _
                     Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "DecodeUTF32LE"
    
    If utf32str = "" Then Exit Function
    
    Dim codepoint As Long
    Dim utf32() As Byte:   utf32 = utf32str
    Dim utf16() As Byte:   ReDim utf16(LBound(utf32) To UBound(utf32))
    Dim i As Long: i = LBound(utf32)
    Dim j As Long: j = i
    
    Do While i < UBound(utf32)
        If utf32(i + 2) = 0 And utf32(i + 3) = 0 Then
            utf16(j) = utf32(i): utf16(j + 1) = utf32(i + 1): j = j + 2
        Else
            If utf32(i + 3) <> 0 Then
                If raiseErrors Then _
                    Err.Raise 5, methodName, _
                    "Codepoint outside of valid Unicode range"
                codepoint = &HFFFD&
            Else
                codepoint = utf32(i + 2) * &H10000 + _
                            utf32(i + 1) * &H100& + utf32(i)
                If codepoint >= &HD800& And codepoint < &HE000& Then
                    If raiseErrors Then _
                        Err.Raise 5, methodName, _
                        "Invalid Unicode codepoint. " & _
                        "(Range reserved for surrogate pairs)"
                    codepoint = &HFFFD&
                ElseIf codepoint > &H10FFFF Then
                    If raiseErrors Then _
                        Err.Raise 5, methodName, _
                        "Codepoint outside of valid Unicode range"
                    codepoint = &HFFFD&
                End If
            End If
            
            Dim n As Long:             n = codepoint - &H10000
            Dim highSurrogate As Long: highSurrogate = &HD800& Or (n \ &H400&)
            Dim lowSurrogate As Long:  lowSurrogate = &HDC00& Or (n And &H3FF)
            
            utf16(j) = highSurrogate And &HFF&
            utf16(j + 1) = highSurrogate \ &H100&
            utf16(j + 2) = lowSurrogate And &HFF&
            utf16(j + 3) = lowSurrogate \ &H100&
            j = j + 4
        End If
        i = i + 4
    Loop
    DecodeUTF32LE = MidB$(utf16, 1, j)
End Function

'Function returning a string containing all alphanumeric characters equally
'distributed. (0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ)
Public Function RandomStringAlphanumeric(ByVal length As Long) As String
    If length < 1 Then Exit Function
    
    Dim i As Long
    Dim char As Long
    Dim b() As Byte: ReDim b(0 To length * 2 - 1)
    
    Randomize
    For i = 0 To length - 1
        Select Case Rnd
            Case Is < 0.41935
                Do: char = 25 * Rnd + 65: Loop Until char <> 0
            Case Is < 0.83871
                Do: char = 25 * Rnd + 97: Loop Until char <> 0
            Case Else
                Do: char = 9 * Rnd + 48: Loop Until char <> 0
        End Select
        
        b(2 * i) = (Int(char)) And &HFF
    Next i
    RandomStringAlphanumeric = b
End Function

'Alternative function returning a string containing all alphanumeric characters
'equally, randomly distributed.
Public Function RandomStringAlphanumeric2(ByVal length As Long) As String
    Static a As Variant
    If IsEmpty(a) Then
        a = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", _
                  "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", _
                  "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", _
                  "w", "x", "y", "z", _
                  "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    End If
    
    Dim i As Long
    Dim result As String: result = Space$(length)
    
    Randomize
    For i = 1 To length
        Mid$(result, i, 1) = a(Int(Rnd() * 62))
    Next i
    RandomStringAlphanumeric2 = result
End Function

'Function returning a string containing all characters from the BMP
'(Basic Multilingual Plane, all 2 byte UTF-16 chars) equally, randomly
'distributed. Excludes surrogate range and BOM.
Public Function RandomStringBMP(ByVal length As Long) As String
    Const MAX_UINT As Long = &HFFFF&
    
    If length < 1 Then Exit Function
    
    Dim i As Long
    Dim char As Long
    Dim b() As Byte:  ReDim b(0 To length * 2 - 1)
    
    Randomize
    For i = 0 To length - 1
        Do
            char = MAX_UINT * Rnd
        Loop Until (char <> 0) _
               And (char < &HD800& Or char > &HDFFF&) _
               And (char <> &HFEFF&)
               
        b(2 * i) = (Int(char)) And &HFF
        b(2 * i + 1) = (Int(char / (&H100))) And &HFF
    Next i
    RandomStringBMP = b
End Function

'Returns a string containing random byte data
Public Function RandomBytes(ByVal numBytes As Long) As String
    Randomize
    Dim bytes() As Byte: ReDim bytes(0 To numBytes - 1)
    Dim i As Long
    For i = 0 To numBytes - 1
        bytes(i) = Rnd * &HFF
    Next i
    RandomBytes = bytes
End Function

'Function returning a string containing all valid unicode characters equally,
'randomly distributed. Excludes surrogate range and BOM.
Public Function RandomStringUnicode(ByVal length As Long) As String
    'Length in UTF-16 codepoints, not unicode codepoints!
    Const MAX_UNICODE As Long = &H10FFFF
    
    If length < 1 Then Exit Function
    
    Dim s As String
    Dim i As Long
    
    Dim char As Long
    Dim b() As Byte: ReDim b(0 To length * 2 - 1)
    
    Randomize
    If length > 1 Then
        For i = 0 To length - 2
            Do
                char = MAX_UNICODE * Rnd
            Loop Until (char <> 0) _
                   And (char < &HD800& Or char > &HDFFF&) _
                   And (char <> &HFEFF&)
                   
            If char < &H10000 Then
                b(2 * i) = (Int(char)) And &HFF
                b(2 * i + 1) = (Int(char / (&H100))) And &HFF
            Else
                Dim m As Long: m = char - &H10000
                Dim highSurrogate As Long: highSurrogate = &HD800& + (m \ &H400&)
                Dim lowSurrogate As Long: lowSurrogate = &HDC00& + (m And &H3FF)
                
                b(2 * i) = highSurrogate And &HFF&
                b(2 * i + 1) = highSurrogate \ &H100&
                i = i + 1
                b(2 * i) = lowSurrogate And &HFF&
                b(2 * i + 1) = lowSurrogate \ &H100&
            End If
        Next i
    End If
    s = b
    If CInt(b(UBound(b) - 1)) + b(UBound(b)) = 0 Then _
        Mid$(s, Len(s), 1) = ChrW(Int(Rnd() * &HFFFE& + 1))
    RandomStringUnicode = s
End Function

'Function returning a string containing all ASCII characters equally,
'randomly distributed.
Public Function RandomStringASCII(ByVal length As Long) As String
    Const MAX_ASC As Long = &H7F&
    Dim i As Long
    Dim char As Integer
    Dim b() As Byte: ReDim b(0 To length * 2 - 1)
    
    Randomize
    For i = 0 To length - 1
        Do: char = MAX_ASC * Rnd: Loop Until char <> 0
        b(2 * i) = (char) And &HFF
    Next i
    RandomStringASCII = b
End Function

'Removes all characters from a string (str) that are not in the string inklChars
'Default inklChars are all alphanumeric characters including dot and space
Public Function CleanString(ByRef str As String, _
                   Optional ByVal inklChars As String = _
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ01234567890. ") _
                            As String
    Dim sChr As String
    Dim i As Long
    Dim j As Long: j = 1
    
    For i = 1 To Len(str)
        sChr = Mid$(str, i, 1)
        
        If InStr(1, inklChars, sChr, vbBinaryCompare) Then
            Mid$(str, j, 1) = sChr
            j = j + 1
        End If
    Next i
    CleanString = Left$(str, j - 1)
End Function

#If Mac = 0 Then
'Removes all non-numeric characters from a string.
'Only keeps codepoints U+0030 - U+0039
Public Function RegExNumOnly(ByRef s As String) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "[^0-9]+"
         RegExNumOnly = .Replace(s, "")
    End With
End Function
#End If

'Removes all non-numeric characters from a string.
'Keeps only codepoints U+0030 - U+0039 AND ALSO
'keeps the Unicode "Fullwidth Digits" (U+FF10 - U+FF19)!
Public Function RemoveNonNumeric(ByVal str As String) As String
    Dim sChr As String
    Dim i As Long
    Dim j As Long: j = 1
    For i = 1 To Len(str)
        sChr = Mid$(str, i, 1)
        If sChr Like "#" Then _
            Mid$(str, j, 1) = sChr: j = j + 1
    Next i
    RemoveNonNumeric = Left$(str, j - 1)
End Function

'Inserts a string into another string at a specified position
'Insert("abcd", "ff", 0) = "ffabcd"
'Insert("abcd", "ff", 1) = "affbcd"
'Insert("abcd", "ff", 3) = "abcffd"
'Insert("abcd", "ff", 4) = "abcdff"
'Insert("abcd", "ff", 9) = "abcdff"
Public Function Insert(ByRef str As String, _
                       ByRef strToInsert As String, _
                       ByRef afterPos As Long) As String
    Const methodName As String = "Insert"
    If afterPos < 0 Then Err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    
    Insert = Mid$(str, 1, afterPos) & strToInsert & Mid$(str, afterPos + 1)
End Function

'Works like Insert but interprets 'afterPos' as byte-index, not char-index
'Inserting at uneven byte positions likely invalidates an utf-16 string!
Public Function InsertB(ByRef str As String, _
                        ByRef strToInsert As String, _
                        ByRef afterPos As Long) As String
    Const methodName As String = "InsertB"
    If afterPos < 0 Then afterPos = 0
    
    InsertB = MidB$(str, 1, afterPos) & strToInsert & MidB$(str, afterPos + 1)
End Function

'Counts the number of times a substring exists in a string. Does not count
'overlapping occurrences of substring.
'E.g.: CountSubstring("abababab", "abab") -> 2
Public Function CountSubstring(ByRef str As String, _
                               ByRef subStr As String, _
                      Optional ByVal lStart As Long = 1, _
                      Optional ByVal lCompare As VbCompareMethod _
                                                 = vbBinaryCompare) As Long
    Const methodName As String = "CountSubstring"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
 
    Dim lenSubStr As Long: lenSubStr = Len(subStr)
    Dim i As Long:         i = InStr(lStart, str, subStr, lCompare)
    
    CountSubstring = 0
    Do Until i = 0
        CountSubstring = CountSubstring + 1
        i = InStr(i + lenSubStr, str, subStr, lCompare)
    Loop
End Function

'Like CountSubstring but scans a string bytewise.
'Example illustrating the difference to CountSubstring:
'                       |c1||c2|
'bytes = HexToString("0x00610061")
'                         |c3|
'sFind =   HexToString("0x6100")
'CountSubstring(bytes, sFind) -> 0
'CountSubstringB(bytes, sFind) -> 1
Public Function CountSubstringB(ByRef bytes As String, _
                                ByRef subStr As String, _
                       Optional ByVal lStart As Long = 1, _
                       Optional ByVal lCompare As VbCompareMethod _
                                               = vbBinaryCompare) As Long
    Const methodName As String = "CountSubstringB"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
      
    Dim lenBSubStr As Long: lenBSubStr = LenB(subStr)
    Dim i As Long:          i = InStrB(lStart, bytes, subStr, lCompare)
    
    CountSubstringB = 0
    Do Until i = 0
        CountSubstringB = CountSubstringB + 1
        i = InStrB(i + lenBSubStr, bytes, subStr, lCompare)
    Loop
End Function

'Works like the inbuilt 'Replace', but parses the string bitwise, not charwise.
'Example illustrating the difference:
'bytes = HexToString("0x00610061")
'sFind = HexToString("0x6100")
'? StringToHex(ReplaceB(bytes, sFind, "")) -> "0x0061"
'? StringToHex(Replace(bytes, sFind, "")) -> "0x00610061"
Public Function ReplaceB(ByRef bytes As String, _
                         ByRef sFind As String, _
                         ByRef sReplace As String, _
                Optional ByVal lStart As Long = 1, _
                Optional ByVal lCount As Long = -1, _
                Optional ByVal lCompare As VbCompareMethod _
                                        = vbBinaryCompare) As String
    Const methodName As String = "ReplaceB"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then Err.Raise 5, methodName, _
        "Argument 'lCount' = " & lCount & " < -1, invalid"
    lCount = lCount And &H7FFFFFFF
    
    If LenB(bytes) = 0 Or LenB(sFind) = 0 Then
        ReplaceB = bytes
        Exit Function
    End If
    
    Dim lenBFind As Long:    lenBFind = LenB(sFind)
    Dim lenBReplace As Long: lenBReplace = LenB(sReplace)
    Dim numRepl As Long:     numRepl = CountSubstringB(bytes, sFind, _
                                                       lStart, lCompare)
    If lCount < numRepl Then numRepl = lCount
    
    Dim buffer() As Byte
    ReDim buffer(0 To LenB(bytes) - lStart + numRepl * (lenBReplace - lenBFind))
    ReplaceB = buffer
    
    Dim i As Long:              i = InStrB(lStart, bytes, sFind, lCompare)
    Dim j As Long:              j = 1
    Dim lastOccurrence As Long: lastOccurrence = lStart
    Dim count As Long:          count = 1
    
    Do Until i = 0 Or count > lCount
        Dim diff As Long: diff = i - lastOccurrence
        MidB$(ReplaceB, j, diff) = MidB$(bytes, lastOccurrence, diff)
        j = j + diff
        If lenBReplace <> 0 Then
            MidB$(ReplaceB, j, lenBReplace) = sReplace
            j = j + lenBReplace
        End If
        count = count + 1
        lastOccurrence = i + lenBFind
        i = InStrB(lastOccurrence, bytes, sFind, lCompare)
    Loop
    MidB$(ReplaceB, j) = MidB$(bytes, lastOccurrence)
End Function

'Replaces repeated occurrences of consecutive 'substring' with a single one
'E.g.: LimitConsecutiveSubstringRepetition("aaaabaaac", "a", 1)  -> "abac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "aa", 1) -> "aabaaac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "a", 2)  -> "aabaac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "ab", 0) -> "aaaaaac"
Public Function LimitConsecutiveSubstringRepetition( _
                                           ByRef str As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Const methodName As String = "LimitConsecutiveSubstringRepetition"
    
    If limit < 0 Then Err.Raise 5, methodName, _
        "Argument 'limit' = " & limit & " < 0, invalid"
    If limit = 0 Then
        LimitConsecutiveSubstringRepetition = Replace(str, subStr, _
                                                      vbNullString, , , compare)
        Exit Function
    Else
        LimitConsecutiveSubstringRepetition = str
    End If
    If Len(str) = 0 Then Exit Function
    If Len(subStr) = 0 Then Exit Function

    Dim i As Long:                i = InStr(1, str, subStr, compare)
    Dim j As Long:                j = 1
    Dim lenSubStr As Long:        lenSubStr = Len(subStr)
    Dim copyChunkSize As Long:    copyChunkSize = 0
    Dim consecutiveCount As Long: consecutiveCount = 0
    Dim lastOccurrence As Long:   lastOccurrence = 1 - lenSubStr
    Dim occurrenceDiff As Long

    Do Until i = 0
        occurrenceDiff = i - lastOccurrence
        If occurrenceDiff = lenSubStr Then
            consecutiveCount = consecutiveCount + 1
            If consecutiveCount <= limit Then
                copyChunkSize = copyChunkSize + occurrenceDiff
            ElseIf consecutiveCount = limit + 1 Then
                Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
                    Mid$(str, i - copyChunkSize, copyChunkSize)
                j = j + copyChunkSize
                copyChunkSize = 0
            End If
        Else
            copyChunkSize = copyChunkSize + occurrenceDiff
            consecutiveCount = 1
        End If
        lastOccurrence = i
        i = InStr(i + lenSubStr, str, subStr, compare)
    Loop

    copyChunkSize = copyChunkSize + Len(str) - lastOccurrence - lenSubStr + 1
    Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
        Mid$(str, Len(str) - copyChunkSize + 1)

    LimitConsecutiveSubstringRepetition = _
        Left$(LimitConsecutiveSubstringRepetition, j + copyChunkSize - 1)
End Function

'Same as LimitConsecutiveSubstringRepetition, but scans the string bytewise.
'Example illustrating the difference:
'Dim bytes As String: bytes = HexToString("0x006100610061")
'Dim subStr As String: subStr = HexToString("0x6100")
'StringToHex(LimitConsecutiveSubstringRepetition(bytes, subStr, 1) _
'    -> "0x006100610061"
'StringToHex(LimitConsecutiveSubstringRepetitionB(bytes, subStr, 1) _
'    -> "0x00610061"
Public Function LimitConsecutiveSubstringRepetitionB( _
                                           ByRef bytes As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Const methodName As String = "LimitConsecutiveSubstringRepetitionB"
    
    If limit < 0 Then Err.Raise 5, methodName, _
        "Argument 'limit' = " & limit & " < 0, invalid"
    If limit = 0 Then
        LimitConsecutiveSubstringRepetitionB = ReplaceB(bytes, subStr, _
                                                      vbNullString, , , compare)
        Exit Function
    Else
        LimitConsecutiveSubstringRepetitionB = bytes
    End If
    If LenB(bytes) = 0 Then Exit Function
    If LenB(subStr) = 0 Then Exit Function

    Dim i As Long:                i = InStrB(1, bytes, subStr, compare)
    Dim j As Long:                j = 1
    Dim lenBSubStr As Long:       lenBSubStr = LenB(subStr)
    Dim copyChunkSize As Long:    copyChunkSize = 0
    Dim consecutiveCount As Long: consecutiveCount = 0
    Dim lastOccurrence As Long:   lastOccurrence = 1 - lenBSubStr
    Dim occurrenceDiff As Long

    Do Until i = 0
        occurrenceDiff = i - lastOccurrence
        If occurrenceDiff = lenBSubStr Then
            consecutiveCount = consecutiveCount + 1
            If consecutiveCount <= limit Then
                copyChunkSize = copyChunkSize + occurrenceDiff
            ElseIf consecutiveCount = limit + 1 Then
                MidB$(LimitConsecutiveSubstringRepetitionB, j, copyChunkSize) = _
                    MidB$(bytes, i - copyChunkSize, copyChunkSize)
                j = j + copyChunkSize
                copyChunkSize = 0
            End If
        Else
            copyChunkSize = copyChunkSize + occurrenceDiff
            consecutiveCount = 1
        End If
        lastOccurrence = i
        i = InStrB(i + lenBSubStr, bytes, subStr, compare)
    Loop

    copyChunkSize = copyChunkSize + LenB(bytes) - lastOccurrence - lenBSubStr + 1
    MidB$(LimitConsecutiveSubstringRepetitionB, j, copyChunkSize) = _
        MidB$(bytes, LenB(bytes) - copyChunkSize + 1)

    LimitConsecutiveSubstringRepetitionB = _
        LeftB$(LimitConsecutiveSubstringRepetitionB, j + copyChunkSize - 1)
End Function

'Repeats the string str, repeatTimes times.
'Works with byte strings of uneven LenB
'E.g.: RepeatString("a", 3) -> "aaa"
'      StrConv(RepeatString(MidB("a", 1, 1), 3), vbUnicode) -> "aaa"
Public Function RepeatString(ByRef str As String, _
                    Optional ByVal repeatTimes As Long = 2) As String
    RepeatString = Space$((LenB(str) * repeatTimes + 1) \ 2)
    
    If (LenB(str) * repeatTimes) Mod 2 = 1 Then _
        RepeatString = MidB$(RepeatString, 1, LenB(RepeatString) - 1)

    Dim i As Long
    For i = 1 To LenB(RepeatString) Step LenB(str)
        MidB$(RepeatString, i, LenB(str)) = str
    Next i
End Function

'Adds fillerStr to the right side of a string repeatedly until the resulting
'string reaches length 'Length'
'E.g.: PadRight("asd", 11, "xyz") -> "asdxyzxyzxy"
Public Function PadRight(ByRef str As String, _
                         ByVal length As Long, _
                Optional ByVal fillerStr As String = " ") As String
    PadRight = PadRightB(str, length * 2, fillerStr)
End Function

'Adds fillerStr to the left side of a string repeatedly until the resulting
'string reaches length 'Length'
'E.g.: PadLeft("asd", 11, "xyz") -> "yzxyzxyzasd"
Public Function PadLeft(ByRef str As String, _
                        ByVal length As Long, _
               Optional ByVal fillerStr As String = " ") As String
    PadLeft = PadLeftB(str, length * 2, fillerStr)
End Function

'Adds fillerStr to the right side of a string repeatedly until the resulting
'string reaches length 'Length' in bytes!
'E.g.: PadRightB("asd", 16, "xyz") -> "asdxyzxy"
Public Function PadRightB(ByRef str As String, _
                          ByVal length As Long, _
                 Optional ByVal fillerStr As String = " ") As String
    Const methodName As String = "PadRightB"
    If length < 0 Then Err.Raise 5, methodName, _
        "Argument 'Length' = " & length & " < 0, invalid"
    If LenB(fillerStr) = 0 Then Err.Raise 5, methodName, _
        "Argument 'fillerStr' = vbNullString, invalid"
    
    If length > LenB(str) Then
        If LenB(fillerStr) = 2 Then
            PadRightB = str & String((length - LenB(str) + 1) \ 2, fillerStr)
            If length Mod 2 = 1 Then _
                PadRightB = LeftB$(PadRightB, LenB(PadRightB) - 1)
        Else
            PadRightB = str & LeftB$(RepeatString(fillerStr, (((length - _
                LenB(str))) + 1) \ LenB(fillerStr) + 1), length - LenB(str))
        End If
    Else
        PadRightB = LeftB$(str, length)
    End If
End Function

'Adds fillerStr to the left side of a string repeatedly until the resulting
'string reaches length 'Length' in bytes!
'Note that this can result in an invalid UTF-16 output for uneven lengths!
'E.g.: PadLeftB("asd", 16, "xyz") -> "yzxyzasd"
'      PadLeftB("asd", 11, "xyz") -> "?????"
Public Function PadLeftB(ByRef str As String, _
                         ByVal length As Long, _
                Optional ByVal fillerStr As String = " ") As String
    Const methodName As String = "PadLeftB"
    If length < 0 Then Err.Raise 5, methodName, _
        "Argument 'Length' = " & length & " < 0, invalid"
    If LenB(fillerStr) = 0 Then Err.Raise 5, methodName, _
        "Argument 'fillerStr' = vbNullString, invalid"
        
    If length > LenB(str) Then
        If LenB(fillerStr) = 2 Then
            PadLeftB = String((length - LenB(str) + 1) \ 2, fillerStr) & str
            If length Mod 2 = 1 Then _
                PadLeftB = RightB$(PadLeftB, LenB(PadLeftB) - 1)
        Else
            PadLeftB = RightB$(RepeatString(fillerStr, (((length - LenB(str))) _
                          + 1) \ LenB(fillerStr) + 1), length - LenB(str)) & str
        End If
    Else
        PadLeftB = RightB$(str, length)
    End If
End Function

'Works like the inbuilt 'Split', but parses string bytewise, so it splits at
'all occurrences of 'Delimiter', even at uneven byte-index positions.
'Example illustrating the difference:
'bytes = HexToString("0x00610061")
'sDelim = HexToString("0x6100")
'SplitB(bytes, sDelim)) -> "0x00", "0x61"
'Split(bytes, sDelim, "")) -> "0x00610061"
Public Function SplitB(ByRef bytes As String, _
              Optional ByRef sDelimiter As String = " ", _
              Optional ByVal lLimit As Long = -1, _
              Optional ByVal lCompare As VbCompareMethod = vbBinaryCompare) _
                       As Variant
    Const methodName As String = "SplitB"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    lLimit = lLimit And &H7FFFFFFF
    
    If LenB(bytes) = 0 Or LenB(sDelimiter) = 0 Or lLimit < 2 Then
        SplitB = VBA.Array(bytes) 'Ignores Option Base 1, like inbuilt 'Split'
        Exit Function
    End If
    
    Dim lenBDelim As Long:  lenBDelim = LenB(sDelimiter)
    Dim numParts As Long:   numParts = CountSubstringB(bytes, sDelimiter, _
                                                       1, lCompare) + 1
    If lLimit < numParts Then numParts = lLimit
    
    Dim arr As Variant:         ReDim arr(0 To numParts - 1)
    Dim i As Long:              i = InStrB(1, bytes, sDelimiter, lCompare)
    Dim lastOccurrence As Long: lastOccurrence = 1
    Dim count As Long:          count = 0

    Do Until i = 0 Or count + 1 >= lLimit
        Dim diff As Long: diff = i - lastOccurrence
        arr(count) = MidB$(bytes, lastOccurrence, diff)
        count = count + 1
        lastOccurrence = i + lenBDelim
        i = InStrB(lastOccurrence, bytes, sDelimiter, lCompare)
    Loop
    arr(count) = MidB$(bytes, lastOccurrence)
    SplitB = arr
End Function

'Splits a string at every occurrence of the specified delimiter "delim", unless
'that delimiter occurs between non-escaped quotes. e.g. (" asf delim asdf ")
'will not be split. Quotes will not be removed.
'Quotes can be excaped by repetition.
'E.g.: SplitUnlessInQuotes("Hello "" ""World" "Goodbye World") returns
'      "Hello "" "" World", and "Goodbye World"
'If " is chosen as delimiter, splits at the outermost two occurrences of ", or
'if only one " exists in the string, splits the string into two parts.
'E.g. SplitUnlessInQuotes("asdf""asdf""asdf""asdf", """") returns
'    "asdf", "asdf""asdf", and "asdf"
Public Function SplitUnlessInQuotes(ByRef str As String, _
                           Optional ByRef delim As String = " ", _
                           Optional limit As Long = -1) As Variant
    Dim i As Long
    Dim s As String
    Dim ub As Long:         ub = -1
    Dim parts As Variant:   ReDim parts(0 To 0)
    Dim doSplit As Boolean: doSplit = True
    
    If delim = """" Then 'Handle this special case
        i = InStr(1, str, """", vbBinaryCompare)
        If i <> 0 Then
            Dim j As Long: j = InStrRev(str, """", , vbBinaryCompare)
            If i = j Then
                SplitUnlessInQuotes = Split(str, """", , vbBinaryCompare)
                Exit Function
            Else
                ReDim parts(0 To 2)
                parts(0) = Left$(str, i - 1)
                parts(1) = Mid$(str, i + 1, j)
                parts(2) = Mid$(str, j + 1)
            End If
        Else
            parts(0) = str
        End If
        SplitUnlessInQuotes = parts
        Exit Function
    End If
    
    For i = 1 To Len(str)
        If ub = limit - 2 Then
            ub = ub + 1
            ReDim Preserve parts(0 To ub)
            parts(ub) = Mid$(str, i)
            Exit For
        End If
        
        If Mid$(str, i, 1) = """" Then
            doSplit = Not doSplit
            If Not doSplit Then _
                doSplit = InStr(i + 1, str, """", vbBinaryCompare) = 0
        End If
        
        If Mid$(str, i, Len(delim)) = delim And doSplit Or i = Len(str) Then
            If i = Len(str) Then s = s & Mid$(str, i, 1)
            ub = ub + 1
            ReDim Preserve parts(0 To ub)
            parts(ub) = s
            s = vbNullString
            i = i + Len(delim) - 1
        Else
            s = s & Mid$(str, i, 1)
        End If
    Next i
    SplitUnlessInQuotes = parts
End Function
