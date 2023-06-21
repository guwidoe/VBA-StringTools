Attribute VB_Name = "LibStringToolsTests"
'===============================================================================
' VBA StringTools - Tests
' ------------------------------------------------------------------------------------
' https://github.com/guwidoe/VBA-StringTools/blob/main/src/test/TestLibStringTools.bas
' ------------------------------------------------------------------------------------
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

Option Explicit

' For source of the timer-code see here:
' https://gist.github.com/guwidoe/5c74c64d79c0e1cd1be458b0632b279a

#If Mac Then
    #If VBA7 Then
        'https://developer.apple.com/documentation/kernel/1462446-mach_absolute_time
        Private Declare PtrSafe Function mach_continuous_time Lib "/usr/lib/libSystem.dylib" () As Currency
        Private Declare PtrSafe Function mach_timebase_info Lib "/usr/lib/libSystem.dylib" (ByRef timebaseInfo As MachTimebaseInfo) As Long
    #Else
        Private Declare Function mach_continuous_time Lib "/usr/lib/libSystem.dylib" () As Currency
        Private Declare Function mach_timebase_info Lib "/usr/lib/libSystem.dylib" (ByRef timebaseInfo As MachTimebaseInfo) As Long
    #End If
#Else
    #If VBA7 Then
        'https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter
        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As LongPtr
        Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As LongPtr
    #Else
        Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef Frequency As Currency) As Long
        Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef Counter As Currency) As Long
    #End If
#End If

#If Mac Then
    Private Type MachTimebaseInfo
        Numerator As Long
        Denominator As Long
    End Type
#End If

'Returns operating system clock tick count since system startup
Private Function GetTickCount() As Currency
    #If Mac Then
        GetTickCount = mach_continuous_time()
    #Else
        QueryPerformanceCounter GetTickCount
    #End If
End Function

'Returns frequency in ticks per second
Private Function GetFrequency() As Currency
    #If Mac Then
        Dim tbInfo As MachTimebaseInfo: mach_timebase_info tbInfo
        
        GetFrequency = (tbInfo.Denominator / tbInfo.Numerator) * 100000@
    #Else
        QueryPerformanceFrequency GetFrequency
    #End If
End Function

'Returns time since system startup in seconds with 0.1ms (=100µs) precision
Private Function AccurateTimer() As Currency
    AccurateTimer = GetTickCount / GetFrequency
End Function

'Returns time since system startup in milliseconds with 0.1µs (=100ns) precision
Private Function AccurateTimerMs() As Currency
    'Note that this calculation will work even if 1000@ / GetFrequency < 0.0001
    AccurateTimerMs = (1000@ / GetFrequency) * GetTickCount
End Function

'Returns time since system startup in microseconds, up to 0.1ns =100ps precision
'The highest precision achieved by this function depends on the system, however,
'typically precision will be the same as for AccurateTimerMs.
Private Function AccurateTimerUs() As Currency
    AccurateTimerUs = (1000000@ / GetFrequency) * GetTickCount
End Function

'###############################################################################
'#########################        UNIT TESTS      ##############################
'###############################################################################
Public Sub RunAllTests()
    TestEncodersAndDecoders
    TestUTF8EncodersPerformance
    TestUTF8DecodersPerformance
    TestUTF32EncodersAndDecodersPerformance
    TestANSIEncodersAndDecodersPerformance
    TestDifferentWaysOfGettingNumericalValuesFromStrings
    RunLimitConsecutiveSubstringRepetitionTests
    RunEscapeUnescapeUnicodeTests
End Sub

Private Sub TestEncodersAndDecoders()
    Const STR_LENGTH As Long = 1000001
    
    Dim fullUnicode As String:    fullUnicode = RandomStringUnicode(STR_LENGTH)
    Dim bmpUnicode As String:     bmpUnicode = RandomStringBMP(STR_LENGTH)
    Dim utf16AsciiOnly As String: utf16AsciiOnly = RandomStringASCII(STR_LENGTH)
    
    'VBA natively implemented Encoders/Decoders
    Debug.Print "Native UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF8native(EncodeUTF8native(bmpUnicode)) = bmpUnicode, "passed", "failed")
        
     Debug.Print "ADODB.Stream UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
         IIf(DecodeUTF8usingAdodbStream(EncodeUTF8usingAdodbStream(bmpUnicode)) = bmpUnicode, "passed", "failed")
         
     Debug.Print "API UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
         IIf(Decode(Encode(bmpUnicode, cpUTF_8), cpUTF_8) = bmpUnicode, "passed", "failed")

    Debug.Print "UTF-32 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF32LE(EncodeUTF32LE(bmpUnicode)) = bmpUnicode, "passed", "failed")
        
    Debug.Print "Native UTF-8 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF8native(EncodeUTF8native(fullUnicode)) = fullUnicode, "passed", "failed")
    
    #If Mac = 0 Then
        Debug.Print "ADODB.Stream UTF-8 Encoder/Decoder Test full Unicode: " & _
            IIf(DecodeUTF8usingAdodbStream(EncodeUTF8usingAdodbStream(fullUnicode)) = fullUnicode, "passed", "failed")
    #End If
    
    Debug.Print "API UTF-8 Encoder/Decoder Test full Unicode: " & _
        IIf(Decode(Encode(fullUnicode, cpUTF_8), cpUTF_8) = fullUnicode, "passed", "failed")
    
    Debug.Print "UTF-32 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF32LE(EncodeUTF32LE(fullUnicode)) = fullUnicode, "passed", "failed")
        
    Debug.Print "ANSI Encoder/Decoder Test: " & _
        IIf(DecodeANSI(EncodeANSI(utf16AsciiOnly)) = utf16AsciiOnly, "passed", "failed")
End Sub

Private Sub TestUTF8EncodersPerformance()
    Dim t As Currency
    
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String
    Dim s As String
    Dim numReps As Long
    Dim strLength As Long
    Dim i As Long
    Dim j As Long
    
    For i = LBound(numRepetitions) To UBound(numRepetitions)
        numReps = numRepetitions(i)
        strLength = strLengths(i)
    
        s = RandomStringUnicode(strLength)
        's = RandomStringBMP(strLength)
        's = RandomStringASCII(strLength)
        
        description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-8 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            EncodeUTF8native s
        Next j
        Debug.Print "EncodeUTF8native took: " & AccurateTimer - t & description
            
        #If Mac = 0 Then
            'ADODB.Stream UTF-8 Encoder:
            t = AccurateTimer
            For j = 1 To numReps
                EncodeUTF8usingAdodbStream s
            Next j
            Debug.Print "EncodeUTF8usingAdodbStream took: " & AccurateTimer - t & description
        #End If
        
        'Windows API UTF-8 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            Encode s, cpUTF_8
        Next j
        Debug.Print "EncodeUTF8usingAPI took: " & AccurateTimer - t & description
        
        DoEvents
    Next i
End Sub

Private Sub TestUTF8DecodersPerformance()
    Dim t As Currency
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String
    Dim s As String
    Dim numReps As Long
    Dim strLength As Long
    Dim i As Long
    Dim j As Long
    
    For i = LBound(numRepetitions) To UBound(numRepetitions)
        numReps = numRepetitions(i)
        strLength = strLengths(i)
    
        s = RandomStringUnicode(strLength)
        's = RandomStringBMP(strLength)
        's = RandomStringASCII(strLength)
        
        s = EncodeUTF8native(s)
        description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-8 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            DecodeUTF8native s
        Next j
        Debug.Print "DecodeUTF8native took: " & AccurateTimer - t & description
        
        #If Mac = 0 Then
            'ADODB.Stream UTF-8 Decoder:
            t = AccurateTimer
            For j = 1 To numReps
                DecodeUTF8usingAdodbStream s
            Next j
            Debug.Print "DecodeUTF8usingAdodbStream took: " & AccurateTimer - t & description
        #End If
        
        'Windows API UTF-8 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            Decode s, cpUTF_8
        Next j
        Debug.Print "DecodeUTF8usingWinAPI took: " & AccurateTimer - t & description
        
        DoEvents
    Next i
End Sub

Private Sub TestUTF32EncodersAndDecodersPerformance()
    Dim t As Currency
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String
    Dim s As String
    Dim s2 As String
    Dim numReps As Long
    Dim strLength As Long
    Dim i As Long
    Dim j As Long
    
    For i = LBound(numRepetitions) To UBound(numRepetitions)
        numReps = numRepetitions(i)
        strLength = strLengths(i)
    
        s = RandomStringUnicode(strLength)
        's = RandomStringBMP(strLength)
        's = RandomStringASCII(strLength)
        
        s2 = EncodeUTF32LE(s)
        description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-32 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            EncodeUTF32LE s
        Next j
        Debug.Print "EncodeUTF32LE took: " & AccurateTimer - t & description
        

        'VBA Native UTF-32 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            DecodeUTF32LE s2
        Next j
        Debug.Print "DecodeUTF32LE took: " & AccurateTimer - t & description

        DoEvents
    Next i
End Sub

Private Sub TestANSIEncodersAndDecodersPerformance()
    Dim t As Currency
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String
    Dim s As String
    Dim s2 As String
    Dim numReps As Long
    Dim strLength As Long
    Dim i As Long
    Dim j As Long
    
    For i = LBound(numRepetitions) To UBound(numRepetitions)
        numReps = numRepetitions(i)
        strLength = strLengths(i)
    
        s = RandomStringUnicode(strLength)
        's = RandomStringBMP(strLength)
        's = RandomStringASCII(strLength)
        
        s2 = EncodeANSI(s)
        description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-32 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            EncodeANSI s
        Next j
        Debug.Print "EncodeANSI took: " & AccurateTimer - t & description
        

        'VBA Native UTF-32 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            DecodeANSI s2
        Next j
        Debug.Print "DecodeANSI took: " & AccurateTimer - t & description

        DoEvents
    Next i
End Sub

Private Sub TestDifferentWaysOfGettingNumericalValuesFromStrings()
    Dim t As Single:   t = Timer()
    Dim str As String: str = RandomStringAlphanumeric(5000000)

    Debug.Print "Creating string took " & Timer - t & " seconds"
    
    t = Timer()
    RemoveNonNumeric str
    Debug.Print "RemoveNonNumeric took " & Timer - t & " seconds"

    t = Timer()
    CleanString str, "0123456789"
    Debug.Print "CleanString took " & Timer - t & " seconds"
    
    #If Mac = 0 Then
        t = Timer()
        RegExNumOnly str
        Debug.Print "RegExNumOnly took " & Timer - t & " seconds"
    #End If
End Sub

Private Sub TestHexToString()
    Dim utf16leTestHexString As String
    utf16leTestHexString = "0x3DD800DE3DD869DC0D203DD869DC3ED8B2DD3DD869DC3DD869DC0D203DD869DC0D203DD867DC0D203DD866DC3ED8B2DD0D203DD869DC0D203DD867DC0D203DD866DC3ED8B2DD0D203DD867DC0D203DD866DC55006E00690063006F006400650053007500700070006F007200740000D800DC6500730074003DD800DE0D203DD869DC3DD869DC0D203DD869DC0D203DD867DC0D203DD866DC3DD881DC3CD8FCDF0D2040260FFE3ED8D4DD3CD8FBDF0D2042260FFE3DD869DC0D2064270FFE0D203DD868DC3CD8C3DF3CD8FBDF0D2040260FFE"
    
    Dim s As String
    s = HexToString(utf16leTestHexString)
    Debug.Print s
End Sub


Public Function LimitConsecutiveSubstringRepetitionCheck(ByVal str As String, _
                                           Optional ByVal subStr As String = vbNewLine, _
                                           Optional ByVal limit As Long = 1, _
                                           Optional ByVal Compare As VbCompareMethod) _
                                                    As String
    Dim sReplace As String:     sReplace = RepeatString(subStr, limit)
    Dim sCompare As String:     sCompare = str
    Do
        Dim sFind As String:    sFind = sReplace & subStr
        Do
            LimitConsecutiveSubstringRepetitionCheck = sCompare
            sCompare = Replace(sCompare, sFind, sReplace, , , Compare)
            sFind = sFind & subStr 'This together with outer loop should
                                   'improve worst-case runtime a lot
        Loop Until sCompare = LimitConsecutiveSubstringRepetitionCheck
    Loop Until sFind = sReplace & subStr & subStr
End Function


Public Function LimitConsecutiveSubstringRepetitionCheck2(ByVal str As String, _
                                           Optional ByVal subStr As String = vbNewLine, _
                                           Optional ByVal limit As Long = 1, _
                                           Optional ByVal Compare As VbCompareMethod) _
                                                    As String
    Dim sReplace As String:     sReplace = RepeatString(subStr, limit)
    Dim sCompare As String:     sCompare = str
    Dim sFind As String:        sFind = sReplace & subStr
    Do
        LimitConsecutiveSubstringRepetitionCheck2 = sCompare
        sCompare = Replace(sCompare, sFind, sReplace, , , Compare)
    Loop Until sCompare = LimitConsecutiveSubstringRepetitionCheck2
End Function

Sub RunLimitConsecutiveSubstringRepetitionTests()
    Dim failedTests As Long
    On Error GoTo errh:
    TestLimitConsecutiveSubstringRepetition "aaaabaaca", "a", 1
    TestLimitConsecutiveSubstringRepetition "aaaabaaca", "aa", 1
    TestLimitConsecutiveSubstringRepetition "abaca", "aa", 1
    TestLimitConsecutiveSubstringRepetition "aaaaabaaca", "aa", 1
    TestLimitConsecutiveSubstringRepetition "aaaaababaca", "ab", 1
    TestLimitConsecutiveSubstringRepetition "bbbaaababbb", "ab", 1
    TestLimitConsecutiveSubstringRepetition _
        UnescapeUnicode("\u6100\u6100\u6100"), "a", 1
    failedTests = failedTests + IIf(LimitConsecutiveSubstringRepetitionB( _
                UnescapeUnicode("\u6100\u6100\u6100"), "a", 1) <> _
            LimitConsecutiveSubstringRepetition( _
                UnescapeUnicode("\u6100\u6100\u6100"), "a", 1), 0, 1)
    'Add more tests here


                    
    If failedTests = 0 Then
        Debug.Print "LimitConsecutiveSubstringRepetition PASSED all tests"
    Else
        Debug.Print "LimitConsecutiveSubstringRepetition FAILED" & failedTests & " tests!"
    End If
    Exit Sub
errh:
    If Err.number = vbObjectError + 43233 Then
        failedTests = failedTests + 1
        Debug.Print Err.description
        Resume Next
    Else
        Err.Raise Err
    End If
End Sub

Private Sub TestLimitConsecutiveSubstringRepetition(ByVal str As String, _
                                  Optional ByVal subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod)
    If LimitConsecutiveSubstringRepetition(str, subStr, limit, Compare) _
    <> LimitConsecutiveSubstringRepetitionCheck(str, subStr, limit, Compare) Then _
        Err.Raise vbObjectError + 43233, "TestLimitConsecutiveSubstringRepetition", _
        "TestLimitConsecutiveSubstringRepetition failed for: " & vbNewLine & _
        "vbCompareMethod: " & Compare & vbNewLine & _
        "limit: " & limit & vbNewLine & _
        "subStr: " & subStr & _
        "str: " & str
End Sub

Sub TestLimitConsecutiveSubstringRepetitionB()
    Dim bytes As String: bytes = HexToString("0x006100610061")
    Dim subStr As String: subStr = HexToString("0x6100")
    Debug.Print StringToHex(LimitConsecutiveSubstringRepetition(bytes, subStr, 1))
    Debug.Print StringToHex(LimitConsecutiveSubstringRepetitionB(bytes, subStr, 0))
End Sub

Private Sub TestReplaceB()
    Dim bytes As String: bytes = HexToString("0x006100610061")
    Dim sFind As String: sFind = HexToString("0x6100")
    Debug.Print "ReplaceB:", StringToHex(ReplaceB(bytes, sFind, ""))
    Debug.Print "Replace:", StringToHex(Replace(bytes, sFind, ""))
End Sub

Private Sub TestSplitB()
    Dim bytes As String: bytes = HexToString("0x006100610061")
    Dim sFind As String: sFind = HexToString("0x6100")
    Dim v As Variant
    v = SplitB(bytes, sFind)
    Debug.Print StringToHex(CStr(v(0))), StringToHex(CStr(v(1))), StringToHex(CStr(v(2)))
    Stop
    v = Split(bytes, sFind)
    Stop
End Sub

Private Static Property Get AllCodePages() As Collection
    Dim C As Collection
    If Not C Is Nothing Then
        Set AllCodePages = C
        Exit Function
    End If
    Set C = New Collection
          'Item: Enum ID, Key:=.NET Name
    C.Add Item:=cpIBM037, Key:="IBM037"
    C.Add Item:=cpIBM437, Key:="IBM437"
    C.Add Item:=cpIBM500, Key:="IBM500"
    C.Add Item:=cpASMO_708, Key:="ASMO-708"
    C.Add Item:=cpASMO_449, Key:="ASMO-449"
    C.Add Item:=cpTransparent_Arabic, Key:="Transparent-Arabic"
    C.Add Item:=cpDOS_720, Key:="DOS-720"
    C.Add Item:=cpIbm737, Key:="ibm737"
    C.Add Item:=cpIbm775, Key:="ibm775"
    C.Add Item:=cpIbm850, Key:="ibm850"
    C.Add Item:=cpIbm852, Key:="ibm852"
    C.Add Item:=cpIBM855, Key:="IBM855"
    C.Add Item:=cpIbm857, Key:="ibm857"
    C.Add Item:=cpIBM00858, Key:="IBM00858"
    C.Add Item:=cpIBM860, Key:="IBM860"
    C.Add Item:=cpIbm861, Key:="ibm861"
    C.Add Item:=cpDOS_862, Key:="DOS-862"
    C.Add Item:=cpIBM863, Key:="IBM863"
    C.Add Item:=cpIBM864, Key:="IBM864"
    C.Add Item:=cpIBM865, Key:="IBM865"
    C.Add Item:=cpCp866, Key:="cp866"
    C.Add Item:=cpIbm869, Key:="ibm869"
    C.Add Item:=cpIBM870, Key:="IBM870"
    C.Add Item:=cpWindows_874, Key:="windows-874"
    C.Add Item:=cpCp875, Key:="cp875"
    C.Add Item:=cpShift_jis, Key:="shift_jis"
    C.Add Item:=cpGb2312, Key:="gb2312"
    C.Add Item:=cpKs_c_5601_1987, Key:="ks_c_5601-1987"
    C.Add Item:=cpBig5, Key:="big5"
    C.Add Item:=cpIBM1026, Key:="IBM1026"
    C.Add Item:=cpIBM01047, Key:="IBM01047"
    C.Add Item:=cpIBM01140, Key:="IBM01140"
    C.Add Item:=cpIBM01141, Key:="IBM01141"
    C.Add Item:=cpIBM01142, Key:="IBM01142"
    C.Add Item:=cpIBM01143, Key:="IBM01143"
    C.Add Item:=cpIBM01144, Key:="IBM01144"
    C.Add Item:=cpIBM01145, Key:="IBM01145"
    C.Add Item:=cpIBM01146, Key:="IBM01146"
    C.Add Item:=cpIBM01147, Key:="IBM01147"
    C.Add Item:=cpIBM01148, Key:="IBM01148"
    C.Add Item:=cpIBM01149, Key:="IBM01149"
    C.Add Item:=cpUTF_16, Key:="utf-16"
    C.Add Item:=cpUnicodeFFFE, Key:="unicodeFFFE"
    C.Add Item:=cpWindows_1250, Key:="windows-1250"
    C.Add Item:=cpWindows_1251, Key:="windows-1251"
    C.Add Item:=cpWindows_1252, Key:="windows-1252"
    C.Add Item:=cpWindows_1253, Key:="windows-1253"
    C.Add Item:=cpWindows_1254, Key:="windows-1254"
    C.Add Item:=cpWindows_1255, Key:="windows-1255"
    C.Add Item:=cpWindows_1256, Key:="windows-1256"
    C.Add Item:=cpWindows_1257, Key:="windows-1257"
    C.Add Item:=cpWindows_1258, Key:="windows-1258"
    C.Add Item:=cpJohab, Key:="Johab"
    C.Add Item:=cpMacintosh, Key:="macintosh"
    C.Add Item:=cpX_mac_japanese, Key:="x-mac-japanese"
    C.Add Item:=cpX_mac_chinesetrad, Key:="x-mac-chinesetrad"
    C.Add Item:=cpX_mac_korean, Key:="x-mac-korean"
    C.Add Item:=cpX_mac_arabic, Key:="x-mac-arabic"
    C.Add Item:=cpX_mac_hebrew, Key:="x-mac-hebrew"
    C.Add Item:=cpX_mac_greek, Key:="x-mac-greek"
    C.Add Item:=cpX_mac_cyrillic, Key:="x-mac-cyrillic"
    C.Add Item:=cpX_mac_chinesesimp, Key:="x-mac-chinesesimp"
    C.Add Item:=cpX_mac_romanian, Key:="x-mac-romanian"
    C.Add Item:=cpX_mac_ukrainian, Key:="x-mac-ukrainian"
    C.Add Item:=cpX_mac_thai, Key:="x-mac-thai"
    C.Add Item:=cpX_mac_ce, Key:="x-mac-ce"
    C.Add Item:=cpX_mac_icelandic, Key:="x-mac-icelandic"
    C.Add Item:=cpX_mac_turkish, Key:="x-mac-turkish"
    C.Add Item:=cpX_mac_croatian, Key:="x-mac-croatian"
    C.Add Item:=cpUTF_32, Key:="utf-32"
    C.Add Item:=cpUTF_32BE, Key:="utf-32BE"
    C.Add Item:=cpX_Chinese_CNS, Key:="x-Chinese_CNS"
    C.Add Item:=cpX_cp20001, Key:="x-cp20001"
    C.Add Item:=cpX_Chinese_Eten, Key:="x_Chinese-Eten"
    C.Add Item:=cpX_cp20003, Key:="x-cp20003"
    C.Add Item:=cpX_cp20004, Key:="x-cp20004"
    C.Add Item:=cpX_cp20005, Key:="x-cp20005"
    C.Add Item:=cpX_IA5, Key:="x-IA5"
    C.Add Item:=cpX_IA5_German, Key:="x-IA5-German"
    C.Add Item:=cpX_IA5_Swedish, Key:="x-IA5-Swedish"
    C.Add Item:=cpX_IA5_Norwegian, Key:="x-IA5-Norwegian"
    C.Add Item:=cpUs_ascii, Key:="us-ascii"
    C.Add Item:=cpX_cp20261, Key:="x-cp20261"
    C.Add Item:=cpX_cp20269, Key:="x-cp20269"
    C.Add Item:=cpIBM273, Key:="IBM273"
    C.Add Item:=cpIBM277, Key:="IBM277"
    C.Add Item:=cpIBM278, Key:="IBM278"
    C.Add Item:=cpIBM280, Key:="IBM280"
    C.Add Item:=cpIBM284, Key:="IBM284"
    C.Add Item:=cpIBM285, Key:="IBM285"
    C.Add Item:=cpIBM290, Key:="IBM290"
    C.Add Item:=cpIBM297, Key:="IBM297"
    C.Add Item:=cpIBM420, Key:="IBM420"
    C.Add Item:=cpIBM423, Key:="IBM423"
    C.Add Item:=cpIBM424, Key:="IBM424"
    C.Add Item:=cpX_EBCDIC_KoreanExtended, Key:="x-EBCDIC-KoreanExtended"
    C.Add Item:=cpIBM_Thai, Key:="IBM-Thai"
    C.Add Item:=cpKoi8_r, Key:="koi8-r"
    C.Add Item:=cpIBM871, Key:="IBM871"
    C.Add Item:=cpIBM880, Key:="IBM880"
    C.Add Item:=cpIBM905, Key:="IBM905"
    C.Add Item:=cpIBM00924, Key:="IBM00924"
    C.Add Item:=cpEuc_jp, Key:="EUC-JP"
    C.Add Item:=cpX_cp20936, Key:="x-cp20936"
    C.Add Item:=cpX_cp20949, Key:="x-cp20949"
    C.Add Item:=cpCp1025, Key:="cp1025"
    C.Add Item:=cpDeprecated, Key:="deprecated"
    C.Add Item:=cpKoi8_u, Key:="koi8-u"
    C.Add Item:=cpIso_8859_1, Key:="iso-8859-1"
    C.Add Item:=cpIso_8859_2, Key:="iso-8859-2"
    C.Add Item:=cpIso_8859_3, Key:="iso-8859-3"
    C.Add Item:=cpIso_8859_4, Key:="iso-8859-4"
    C.Add Item:=cpIso_8859_5, Key:="iso-8859-5"
    C.Add Item:=cpIso_8859_6, Key:="iso-8859-6"
    C.Add Item:=cpIso_8859_7, Key:="iso-8859-7"
    C.Add Item:=cpIso_8859_8, Key:="iso-8859-8"
    C.Add Item:=cpIso_8859_9, Key:="iso-8859-9"
    C.Add Item:=cpIso_8859_13, Key:="iso-8859-13"
    C.Add Item:=cpIso_8859_15, Key:="iso-8859-15"
    C.Add Item:=cpX_Europa, Key:="x-Europa"
    C.Add Item:=cpIso_8859_8_i, Key:="iso-8859-8-i"
    C.Add Item:=cpIso_2022_jp, Key:="iso-2022-jp"
    C.Add Item:=cpCsISO2022JP, Key:="csISO2022JP"
    C.Add Item:=cpIso_2022_jp_w_1b_Kana, Key:="iso-2022-jp_w-1b-Kana"
    C.Add Item:=cpIso_2022_kr, Key:="iso-2022-kr"
    C.Add Item:=cpX_cp50227, Key:="x-cp50227"
    C.Add Item:=cpISO_2022_Trad_Chinese, Key:="ISO-2022-Traditional-Chinese"
    C.Add Item:=cpEBCDIC_Jap_Katakana_Ext, Key:="EBCDIC-Japanese-Katakana-Extended"
    C.Add Item:=cpEBCDIC_US_Can_and_Jap, Key:="EBCDIC-US-Canada-and-Japanese"
    C.Add Item:=cpEBCDIC_Kor_Ext_and_Kor, Key:="EBCDIC-Korean-Extended-and-Korean"
    C.Add Item:=cpEBCDIC_Simp_Chin_Ext, Key:="EBCDIC-Simplified-Chinese-Extended-and-Simplified-Chinese"
    C.Add Item:=cpEBCDIC_Simp_Chin, Key:="EBCDIC-Simplified-Chinese"
    C.Add Item:=cpEBCDIC_US_Can_Trad_Chin, Key:="EBCDIC-US-Canada-and-Traditional-Chinese"
    C.Add Item:=cpEBCDIC_Jap_Latin_Ext, Key:="EBCDIC-Japanese-Latin-Extended-and-Japaneseeuc_jp"
    C.Add Item:=cpEUC_CN, Key:="EUC-CN"
    C.Add Item:=cpEuc_kr, Key:="euc-kr"
    C.Add Item:=cpEUC_Traditional_Chinese, Key:="EUC-Traditional-Chinese"
    C.Add Item:=cpHz_gb_2312, Key:="hz-gb-2312"
    C.Add Item:=cpGB18030, Key:="GB18030"
    C.Add Item:=cpX_iscii_de, Key:="x-iscii-de"
    C.Add Item:=cpX_iscii_be, Key:="x-iscii-be"
    C.Add Item:=cpX_iscii_ta, Key:="x-iscii-ta"
    C.Add Item:=cpX_iscii_te, Key:="x-iscii-te"
    C.Add Item:=cpX_iscii_as, Key:="x-iscii-as"
    C.Add Item:=cpX_iscii_or, Key:="x-iscii-or"
    C.Add Item:=cpX_iscii_ka, Key:="x-iscii-ka"
    C.Add Item:=cpX_iscii_ma, Key:="x-iscii-ma"
    C.Add Item:=cpX_iscii_gu, Key:="x-iscii-gu"
    C.Add Item:=cpX_iscii_pa, Key:="x-iscii-pa"
    C.Add Item:=cpUTF_7, Key:="utf-7"
    C.Add Item:=cpUTF_8, Key:="utf-8"

    Set AllCodePages = C
End Property

Sub TestAPI()
    Dim i As Long
    Dim cpID As Variant
    Dim rndBytes As String
    rndBytes = RandomStringUnicode(1000)
    Dim convNotSupported() As Boolean
    ReDim convNotSupported(1 To 151)
    On Error Resume Next
    For Each cpID In AllCodePages
        Encode rndBytes, cpID, True
        i = i + 1
        Debug.Print i, cpID, Err.number, Err.description
        convNotSupported(i) = Err.number
        On Error GoTo -1
    Next cpID
'    i = 0
'    For Each cpID In AllCodePages
'        Encode rndBytes, cpID, False, False
'        i = i + 1
'        'Debug.Print i, cpID, Err.Number, Err.description
'        If (convNotSupported(i) = False) And Err.Number <> 0 Then
'            Debug.Print i, cpID, Err.Number, Err.description
'        End If
'        On Error GoTo -1
'    Next cpID
End Sub

Private Sub RunEscapeUnescapeUnicodeTests()
    TestUnicodeFunctionality
    TestEscapeAndUnescapeUnicode
    TestEscapeUnescapeUnicodePerformance
End Sub

Private Sub TestUnicodeFunctionality()
    Dim originalStr As String
    Dim escapedStr As String
    Dim unescapedStr As String
    Dim formatTypes As UnicodeEscapeFormat
    Dim result As Boolean
    Dim i As Integer
    
    'Test for all types of UnicodeEscapeFormat
    For i = 1 To efAll
        formatTypes = i
        
        'Generate a random string for testing
        originalStr = RandomStringUnicode(10000)
        
        'Test the EscapeUnicode function
        escapedStr = EscapeUnicode(originalStr, &HFF, formatTypes)
        
        'Test the UnescapeUnicode function
        unescapedStr = UnescapeUnicode(escapedStr, formatTypes)
        
        'Check if the unescaped string is equal to the original string
        If originalStr <> unescapedStr Then
            Debug.Print "FAILED Escape/UnescapeUnicode Test for format " & formatTypes
        Else
            Debug.Print "PASSED Escape/UnescapeUnicode Test for format " & formatTypes
        End If
    Next i
End Sub

Private Sub TestEscapeAndUnescapeUnicode()
    Dim originalStr As String
    Dim escapedStr As String
    Dim unescapedStr As String
    Dim formatTypes As UnicodeEscapeFormat
    Dim result As Boolean
    Dim i As Long

    For i = 1 To 100000
        'Create any random combination of formats excluding efUPlus, because
        'efUPlus has a high likelyhood of creating strings that will convert
        'back to a different string than the original string
        Do Until formatTypes <> 0
            formatTypes = Int(15 * Rnd) + 1
            formatTypes = formatTypes And (&HFFFFFFFF - efUPlus)
        Loop
        
        'Generate a random string for testing
        Select Case i Mod 4
            Case 0
                originalStr = RandomStringASCII(10)
            Case 1
                originalStr = RandomStringAlphanumeric(10)
            Case 2
                originalStr = RandomStringBMP(10)
            Case 3
                originalStr = RandomStringUnicode(10)
        End Select
    
        'Test the EscapeUnicode function
        escapedStr = EscapeUnicode(originalStr, i Mod 127, formatTypes)

        'Test the UnescapeUnicode function
        unescapedStr = UnescapeUnicode(escapedStr, formatTypes)

        'Check if the unescaped string is equal to the original string
        If originalStr <> unescapedStr Then
            Debug.Print i
            Debug.Print "originalStr", originalStr
            Debug.Print "escapedStr", escapedStr
            Debug.Print "unescapedStr", unescapedStr
            Debug.Print StringToHex(originalStr)
            Debug.Print StringToHex(unescapedStr)
            Debug.Print "FAILED Escape/UnescapeUnicode Test!"
            Exit Sub
        End If
    Next i
    Debug.Print "PASSED Escape/UnescapeUnicode Stress Test!"
End Sub

Private Sub TestEscapeUnescapeUnicodePerformance()
    Dim originalStr As String
    Dim escapedStr As String
    Dim unescapedStr As String
    Dim formatTypes As UnicodeEscapeFormat
    Dim startTime As Currency
    Dim endTime As Currency
    Dim elapsedTime As Currency
    Dim i As Integer
    
    'Test for all types of UnicodeEscapeFormat
    For i = 1 To efAll '(Log(efAll + 1) / Log(2)) - 1
        formatTypes = i '2 ^ i
        
        'Generate a large random string for testing
        originalStr = RandomStringUnicode(100000)
        
        'Start the timer
        startTime = AccurateTimerMs()
        
        'Test the EscapeUnicode function
        escapedStr = EscapeUnicode(originalStr, , formatTypes)
        
        Debug.Print "UnicodeEscapeFormat: " & formatTypes & _
                    " Escaping took: ", AccurateTimerMs() - startTime & " ms"
        
        startTime = AccurateTimerMs()
        
        'Test the UnescapeUnicode function
        unescapedStr = UnescapeUnicode(escapedStr, formatTypes)
        
        Debug.Print "UnicodeEscapeFormat: " & formatTypes & _
                    " Unescaping took: ", AccurateTimerMs() - startTime & " ms"
    Next i
End Sub

Sub TestReplaceMultiple()

    Dim s As String
    s = RandomStringAlphanumeric(100000)
    Dim finds As Variant
    finds = StringToCodepointStrings(RandomStringUnicode(10000)) 'VBA.Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
    Dim replaces As Variant
     replaces = StringToCodepointStrings(RandomStringUnicode(5000)) ' Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j")
    'Debug.Print ReplaceMultiple(s, finds, replaces) = ReplaceMultipleMultiPass(s, finds, replaces)
    'Debug.Print ReplaceMultiple(s, finds, replaces)
    'Debug.Print ReplaceMultipleMultiPass(s, finds, replaces)
    st
    ReplaceMultiple s, finds, replaces
    RT "ReplaceMultiple", , True
    ReplaceMultipleMultiPass s, finds, replaces
    RT "ReplaceMultipleMultiPass"
    
    'Debug.Print ReplaceMultipleB(s, Array("1", "2", "3"), Array("44", "55"))
End Sub

Public Sub TestDebugPrintArray()
    
    ' Test Case 1: Single dimensional array of integers
    Dim array1DInt(1 To 100) As Integer
    Dim i As Long
    For i = 1 To 100
        array1DInt(i) = i
    Next i
    Debug.Print "Test Case 1: Single dimensional array of integers"
    DebugPrintArray array1DInt
    
    ' Test Case 2: Single dimensional array of strings
    Dim array1DStr(1 To 3) As String
    array1DStr(1) = "Alice"
    array1DStr(2) = "Bob"
    array1DStr(3) = "Charlie"
    Debug.Print "Test Case 2: Single dimensional array of strings"
    DebugPrintArray array1DStr
    
    ' Test Case 3: Two dimensional array of integers
    Dim array2DInt() As Long
    ReDim array2DInt(1 To 100, 1 To 100)
    Dim j As Integer
    For i = 1 To 100
        For j = 1 To 100
            array2DInt(i, j) = i * j
        Next j
    Next i
    Debug.Print "Test Case 3: Two dimensional array of integers"
    DebugPrintArray array2DInt

    Dim array2DStr() As String
    ReDim array2DStr(1 To 100, 1 To 100)
    For i = 1 To 100
        For j = 1 To 100
            array2DStr(i, j) = RandomStringAlphanumeric(Rnd * 15)
        Next j
    Next i
    Debug.Print "Test Case 3: Two dimensional array of Strings"
    DebugPrintArray array2DStr

    ' Test Case 4: Two dimensional array of strings
    ReDim array2DStr(1 To 2, 1 To 2) As String
    array2DStr(1, 1) = "Apple"
    array2DStr(1, 2) = "Banana"
    array2DStr(2, 1) = "Cherry"
    array2DStr(2, 2) = "Durian"
    Debug.Print "Test Case 4: Two dimensional array of strings"
    DebugPrintArray array2DStr
    
    ' Test Case 5: Array containing random strings with special characters
    Dim arrayRandom(1 To 3) As String
    arrayRandom(1) = RandomString(10, 33, 127) ' printable ASCII characters
    arrayRandom(2) = RandomString(10, 256, 500) ' extended ASCII characters
    arrayRandom(3) = RandomString(10, &H1F600, &H1F64F) ' emojis
    Debug.Print "Test Case 5: Array containing random strings with special characters"
    DebugPrintArray arrayRandom, escapeNonPrintableCodepoints:=False
    DebugPrintArray arrayRandom, escapeNonPrintableCodepoints:=True
    
    ' Test Case 6: Empty array
    Dim emptyArray() As Integer
    Debug.Print "Test Case 6: Empty array"
    DebugPrintArray emptyArray

End Sub

Public Sub Array2DToImmediate(ByVal arr As Variant, _
                                Optional ByVal spaces_between_columns As Long = 2, _
                                Optional ByVal NrOfColsToOutlineLeft As Long = 2)
'Prints a 2D-array of values to a table (with same sized column widhts) in the immmediate window

'Each character in the Immediate window of VB Editor (CTRL+G to show) has the same pixel width,
'thus giving the option to output a proper looking 2D-array (a table with variable string lenghts).
Dim i As Long, j As Long
Dim arrMaxLenPerCol() As Long
Dim str As String
Dim maxLength As Long: maxLength = 198 * 1021& 'capacity of Immediate window is about 200 lines of 1021 characters per line.

'determine max stringlength per column
ReDim arrMaxLenPerCol(UBound(arr, 1))
For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(arr, 2) To UBound(arr, 2)
        arrMaxLenPerCol(i) = IIf(Len(arr(i, j)) > arrMaxLenPerCol(i), Len(arr(i, j)), arrMaxLenPerCol(i))
    Next j
Next i

'build table
For j = LBound(arr, 2) To UBound(arr, 2)
    For i = LBound(arr, 1) To UBound(arr, 1)
        'outline left --> value & spaces & column_spaces
        If i < NrOfColsToOutlineLeft Then
            On Error Resume Next
            str = str & arr(i, j) & Space$((arrMaxLenPerCol(i) - Len(arr(i, j)) + spaces_between_columns) * 1)
        
        'last column to outline left --> value & spaces
        ElseIf i = NrOfColsToOutlineLeft Then
            On Error Resume Next
            str = str & arr(i, j) & Space$((arrMaxLenPerCol(i) - Len(arr(i, j))) * 1)
                    
        'outline right --> spaces & column_spaces & value
        Else 'i > NrOfColsToOutlineLeft Then
            On Error Resume Next
            str = str & Space$((arrMaxLenPerCol(i) - Len(arr(i, j)) + spaces_between_columns) * 1) & arr(i, j)
        End If
    Next i
    str = str & vbNewLine
    If Len(str) > maxLength Then GoTo theEnd
Next j

theEnd:
'capacity of Immediate window is about 200 lines of 1021 characters per line.
If Len(str) > maxLength Then str = Left(str, maxLength) & vbNewLine & " - Table to large for Immediate window"
Debug.Print str
End Sub
