Attribute VB_Name = "LibStringToolsTests"
'===============================================================================
' VBA StringTools - Tests
' ------------------------------------------------------------------------------------
' https://github.com/guwidoe/VBA-StringTools/blob/main/src/test/TestLibStringTools.bas
' ------------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2023 Guido Witt-Dörring
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

#If VBA7 Then
    Private Declare PtrSafe Function IsValidCodePage Lib "kernel32" (ByVal codePage As Long) As Long
#Else
    Private Declare Function IsValidCodePage Lib "kernel32" (ByVal CodePage As Long) As Long
#End If

Private Type CpInfo 'Custom extended CpInfo type for use in this library
    'From CpInfoExW:
    codePage As Long              ' code page id
    MaxCharSize As Long           ' max length (in bytes) of a character
    defaultChar As String         ' default character (MB)
    LeadByte As String            ' lead byte ranges
    UnicodeDefaultChar As String  ' default character (Unicode)
    CodePageName As String        ' code page name (Unicode)
    'Extra:
    AllowsFlags As Boolean
    AllowsQueryReversible As Boolean
    MacConvDescriptorName As String
    IsInitialized As Boolean
End Type

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
'Todo: the 'RunAllTests' currently only runs a fraction of the tests, this
'module needs work to make it consistent and tidy
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
    Const ADD_BOM As Boolean = False
    Const ResPass As String = "Ok"
    Const ResFail As String = "----FAILED----"
    
    Dim bom As String: If ADD_BOM Then bom = ChrU(&HFEFF&)
    Dim fullUnicode As String:     fullUnicode = bom & RandomStringUnicode(STR_LENGTH, True)
    Dim fullUnicodeUTF8 As String: fullUnicodeUTF8 = Encode(fullUnicode, cpUTF_8)
    Dim bmpUnicode As String:      bmpUnicode = bom & RandomStringBMP(STR_LENGTH, True)
    Dim utf16AsciiOnly As String:  utf16AsciiOnly = RandomStringASCII(STR_LENGTH, True)
    Dim rndBytes As String:        rndBytes = bom & RandomBytes(STR_LENGTH, True)
    
    'VBA natively implemented Encoders/Decoders
    Debug.Print "Native UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF8(EncodeUTF8(bmpUnicode)) = bmpUnicode, ResPass, ResFail)
       
    #If Mac = 0 Then
        Debug.Print "ADODB.Stream UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
             IIf(DecodeUTF8usingAdodbStream(EncodeUTF8usingAdodbStream(bmpUnicode)) = bmpUnicode, ResPass, ResFail)
    #End If
    
    Debug.Print "API UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
         IIf(Decode(Encode(bmpUnicode, cpUTF_8), cpUTF_8) = bmpUnicode, ResPass, ResFail)

    Debug.Print "UTF-32 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF32LE(EncodeUTF32LE(bmpUnicode)) = bmpUnicode, ResPass, ResFail)
        
    Debug.Print "Native UTF-8 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF8(EncodeUTF8(fullUnicode)) = fullUnicode, ResPass, ResFail)
    
    Debug.Print "Native UTF-8 Encoder vs API UTF-8 Encoder Test full Unicode: " & _
        IIf(EncodeUTF8(fullUnicode) = Encode(fullUnicode, cpUTF_8), ResPass, ResFail)
        
    Debug.Print "Native UTF-8 Decoder vs API UTF-8 Decoder Test full Unicode: " & _
        IIf(DecodeUTF8(fullUnicodeUTF8) = Decode(fullUnicodeUTF8, cpUTF_8), ResPass, ResFail)
        
    Debug.Print "Native UTF-8 Encoder vs API UTF-8 Encoder Test rndBytes: " & _
        IIf(EncodeUTF8(rndBytes) = Encode(rndBytes, cpUTF_8), ResPass, ResFail)
        
    Debug.Print "Native UTF-8 Decoder vs API UTF-8 Decoder Test rndBytes: " & _
        IIf(DecodeUTF8(rndBytes) = Decode(rndBytes, cpUTF_8), ResPass, ResFail)
    
    #If Mac = 0 Then
        Debug.Print "ADODB.Stream UTF-8 Encoder/Decoder Test full Unicode: " & _
            IIf(DecodeUTF8usingAdodbStream(EncodeUTF8usingAdodbStream(fullUnicode)) = fullUnicode, ResPass, ResFail)
    #End If
    
    Debug.Print "API UTF-8 Encoder/Decoder Test full Unicode: " & _
        IIf(Decode(Encode(fullUnicode, cpUTF_8), cpUTF_8) = fullUnicode, ResPass, ResFail)
    
    Debug.Print "UTF-32 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF32LE(EncodeUTF32LE(fullUnicode)) = fullUnicode, ResPass, ResFail)
        
    Debug.Print "ANSI Encoder/Decoder Test: " & _
        IIf(DecodeANSI(EncodeANSI(utf16AsciiOnly)) = utf16AsciiOnly, ResPass, ResFail)
End Sub

Private Sub TestUTF8EncodersPerformance()
    Dim t As Currency
    
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim Description As String
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
        
        Description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-8 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            EncodeUTF8 s
        Next j
        Debug.Print "EncodeUTF8 took: " & AccurateTimer - t & Description
            
        #If Mac = 0 Then
            'ADODB.Stream UTF-8 Encoder:
            t = AccurateTimer
            For j = 1 To numReps
                EncodeUTF8usingAdodbStream s
            Next j
            Debug.Print "EncodeUTF8usingAdodbStream took: " & AccurateTimer - t & Description
        #End If
        
        'Windows API UTF-8 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            Encode s, cpUTF_8
        Next j
        Debug.Print "EncodeUTF8usingAPI took: " & AccurateTimer - t & Description
        
        DoEvents
    Next i
End Sub

Private Sub TestUTF8DecodersPerformance()
    Dim t As Currency
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim Description As String
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
        
        s = EncodeUTF8(s)
        Description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-8 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            DecodeUTF8 s
        Next j
        Debug.Print "DecodeUTF8native took: " & AccurateTimer - t & Description
        
        #If Mac = 0 Then
            'ADODB.Stream UTF-8 Decoder:
            t = AccurateTimer
            For j = 1 To numReps
                DecodeUTF8usingAdodbStream s
            Next j
            Debug.Print "DecodeUTF8usingAdodbStream took: " & AccurateTimer - t & Description
        #End If
        
        'Windows API UTF-8 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            Decode s, cpUTF_8
        Next j
        Debug.Print "DecodeUTF8usingWinAPI took: " & AccurateTimer - t & Description
        
        DoEvents
    Next i
End Sub

Private Sub TestUTF32EncodersAndDecodersPerformance()
    Dim t As Currency
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim Description As String
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
        Description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-32 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            EncodeUTF32LE s
        Next j
        Debug.Print "EncodeUTF32LE took: " & AccurateTimer - t & Description
        

        'VBA Native UTF-32 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            DecodeUTF32LE s2
        Next j
        Debug.Print "DecodeUTF32LE took: " & AccurateTimer - t & Description

        DoEvents
    Next i
End Sub

Private Sub TestANSIEncodersAndDecodersPerformance()
    Dim t As Currency
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant: numRepetitions = VBA.Array(100000, 1000, 10)
    Dim strLengths As Variant:     strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim Description As String
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
        Description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-32 Encoder:
        t = AccurateTimer
        For j = 1 To numReps
            EncodeANSI s
        Next j
        Debug.Print "EncodeANSI took: " & AccurateTimer - t & Description
        

        'VBA Native UTF-32 Decoder:
        t = AccurateTimer
        For j = 1 To numReps
            DecodeANSI s2
        Next j
        Debug.Print "DecodeANSI took: " & AccurateTimer - t & Description

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




Private Function TestLimitConsecutiveSubstringRepetition(ByVal str As String, _
                                  Optional ByVal subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod)
    On Error Resume Next
    TestLimitConsecutiveSubstringRepetition = "passed"
    Dim res1 As String
    res1 = LimitConsecutiveSubstringRepetition(str, subStr, limit, Compare)
    Dim res2 As String
    res2 = LimitConsecutiveSubstringRepetitionCheck(str, subStr, limit, Compare)
    On Error GoTo 0
    
    If res1 = res2 Then Exit Function
    TestLimitConsecutiveSubstringRepetition = "failed"
    
    Err.Raise vbObjectError + 43233, "TestLimitConsecutiveSubstringRepetition", _
        "TestLimitConsecutiveSubstringRepetition failed for: " & vbNewLine & _
        "vbCompareMethod: " & Compare & vbNewLine & _
        "limit: " & limit & vbNewLine & _
        "subStr: '" & subStr & "'" & vbNewLine & _
        "str: '" & str & "'"
End Function

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
    'Debug.Print "ReplaceFastB:", StringToHex(ReplaceFastB(bytes, sFind, ""))
    Debug.Print "Replace:", StringToHex(Replace(bytes, sFind, ""))
End Sub

'Working accurate version of ReplaceB to compare results against for random
'input testing
Public Function ReplaceBCheck(ByRef bytes As String, _
                         ByRef sFind As String, _
                         ByRef sReplace As String, _
                Optional ByVal lStart As Long = 1, _
                Optional ByVal lCount As Long = -1, _
                Optional ByVal lCompare As VbCompareMethod _
                                        = vbBinaryCompare) As String
    Const methodName As String = "ReplaceBCheck"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then Err.Raise 5, methodName, _
        "Argument 'lCount' = " & lCount & " < -1, invalid"
    lCount = lCount And &H7FFFFFFF

    If LenB(bytes) = 0 Or LenB(sFind) = 0 Then
        ReplaceBCheck = MidB$(bytes, lStart)
        Exit Function
    End If

    Dim lenBFind As Long:         lenBFind = LenB(sFind)
    Dim lenBReplace As Long:      lenBReplace = LenB(sReplace)
    Dim bufferSizeChange As Long
    bufferSizeChange = CountSubstringB(bytes, sFind, lStart, lCount, lCompare) _
                                             * (lenBReplace - lenBFind) - lStart
    
    If LenB(bytes) + bufferSizeChange < 0 Then Exit Function
    
    Dim buffer() As Byte: ReDim buffer(0 To LenB(bytes) + bufferSizeChange)
    ReplaceBCheck = buffer

    Dim i As Long:              i = InStrB(lStart, bytes, sFind, lCompare)
    Dim j As Long:              j = 1
    Dim lastOccurrence As Long: lastOccurrence = lStart
    Dim count As Long:          count = 1

    Do Until i = 0 Or count > lCount
        Dim diff As Long: diff = i - lastOccurrence
        If diff > 0 Then _
            MidB$(ReplaceBCheck, j, diff) = MidB$(bytes, lastOccurrence, diff)
        j = j + diff
        If lenBReplace <> 0 Then
            MidB$(ReplaceBCheck, j, lenBReplace) = sReplace
            j = j + lenBReplace
        End If
        count = count + 1
        lastOccurrence = i + lenBFind
        i = InStrB(lastOccurrence, bytes, sFind, lCompare)
    Loop
    If j <= LenB(ReplaceBCheck) Then MidB$(ReplaceBCheck, j) = MidB$(bytes, lastOccurrence)
End Function

'Sub for ReplaceFast functionality testing
Private Sub TestReplaceFast()
    Dim s As String, f As String, r As String
    Dim st As Long:  st = 1 'default
    Dim c As Long:   c = -1 'default
    Dim cmp As VbCompareMethod: cmp = vbBinaryCompare
    s = "abcde" & Space(10000) & "fghijk"
    f = "fg"
    r = "asdfadsf"
    Debug.Print ReplaceFast(s, f, r, st, c, cmp) = Replace(s, f, r, st, c, cmp)
    s = "fgsafas" & Space(10000) & "dadsfg"
    Debug.Print ReplaceFast(s, f, r, st, c, cmp) = Replace(s, f, r, st, c, cmp)
    s = "fgfgfgfgfgfg" & RepeatString("fg", 12000) & "fgfgasdfasdffgfgfg"
    c = 11000
    Debug.Print ReplaceFast(s, f, r, st, c, cmp) = Replace(s, f, r, st, c, cmp)
    st = 4
    Debug.Print ReplaceFast(s, f, r, st, c, cmp) = Replace(s, f, r, st, c, cmp)
End Sub

''Placeholder for potential dev ReplaceFastB version for comparing accuracy with
''already implemented version
'Private Sub TestReplaceFastB()
'    Dim s As String, f As String, r As String
'    Dim st As Long:  st = 1 'default
'    Dim c As Long:   c = -1 'default
'    Dim cmp As VbCompareMethod: cmp = vbBinaryCompare
'    s = PadRightB(" ", 1) & "abcde" & Space(10000) & "fghijk"
'    f = "fg"
'    r = PadRightB(" ", 1) & "asdfadsf"
'    Debug.Print ReplaceFastB(s, f, r, st, c, cmp) = ReplaceB(s, f, r, st, c, cmp)
'    s = "fgsafas" & Space(10000) & "dadsfg"
'    Debug.Print ReplaceFastB(s, f, r, st, c, cmp) = ReplaceB(s, f, r, st, c, cmp)
'    s = "fgfgfgfgfgfg" & RepeatString("fg", 12000) & "fgfgasdfasdffgfgfg"
'    c = 11000
'    Debug.Print ReplaceFastB(s, f, r, st, c, cmp) = ReplaceB(s, f, r, st, c, cmp)
'    st = 4
'    Debug.Print ReplaceFastB(s, f, r, st, c, cmp) = ReplaceB(s, f, r, st, c, cmp)
'End Sub

Private Sub TestSplitB()
    Dim bytes As String: bytes = HexToString("0x006100610061")
    Dim sFind As String: sFind = HexToString("0x6100")
    Dim v As Variant
    v = SplitB(bytes, sFind)
    Debug.Print StringToHex(CStr(v(0))), StringToHex(CStr(v(1))), StringToHex(CStr(v(2)))

    v = Split(bytes, sFind)

    'v = SplitB(bytes, sFind, 0)
    v = Split(bytes, , 0)
    Stop
End Sub

Private Static Property Get AllCodePages() As Collection
    Dim c As Collection
    If Not c Is Nothing Then
        Set AllCodePages = c
        Exit Property
    End If
    Set c = New Collection
          'Item: Enum ID, Key:=.NET Name
    c.Add Item:=cpIBM037, Key:="IBM037"
    c.Add Item:=cpIBM437, Key:="IBM437"
    c.Add Item:=cpIBM500, Key:="IBM500"
    c.Add Item:=cpASMO_708, Key:="ASMO-708"
    c.Add Item:=cpASMO_449, Key:="ASMO-449"
    c.Add Item:=cpTransparent_Arabic, Key:="Transparent-Arabic"
    c.Add Item:=cpDOS_720, Key:="DOS-720"
    c.Add Item:=cpIbm737, Key:="ibm737"
    c.Add Item:=cpIbm775, Key:="ibm775"
    c.Add Item:=cpIbm850, Key:="ibm850"
    c.Add Item:=cpIbm852, Key:="ibm852"
    c.Add Item:=cpIBM855, Key:="IBM855"
    c.Add Item:=cpIbm857, Key:="ibm857"
    c.Add Item:=cpIBM00858, Key:="IBM00858"
    c.Add Item:=cpIBM860, Key:="IBM860"
    c.Add Item:=cpIbm861, Key:="ibm861"
    c.Add Item:=cpDOS_862, Key:="DOS-862"
    c.Add Item:=cpIBM863, Key:="IBM863"
    c.Add Item:=cpIBM864, Key:="IBM864"
    c.Add Item:=cpIBM865, Key:="IBM865"
    c.Add Item:=cpCp866, Key:="cp866"
    c.Add Item:=cpIbm869, Key:="ibm869"
    c.Add Item:=cpIBM870, Key:="IBM870"
    c.Add Item:=cpWindows_874, Key:="windows-874"
    c.Add Item:=cpCp875, Key:="cp875"
    c.Add Item:=cpShift_jis, Key:="shift_jis"
    c.Add Item:=cpGb2312, Key:="gb2312"
    c.Add Item:=cpKs_c_5601_1987, Key:="ks_c_5601-1987"
    c.Add Item:=cpBig5, Key:="big5"
    c.Add Item:=cpIBM1026, Key:="IBM1026"
    c.Add Item:=cpIBM01047, Key:="IBM01047"
    c.Add Item:=cpIBM01140, Key:="IBM01140"
    c.Add Item:=cpIBM01141, Key:="IBM01141"
    c.Add Item:=cpIBM01142, Key:="IBM01142"
    c.Add Item:=cpIBM01143, Key:="IBM01143"
    c.Add Item:=cpIBM01144, Key:="IBM01144"
    c.Add Item:=cpIBM01145, Key:="IBM01145"
    c.Add Item:=cpIBM01146, Key:="IBM01146"
    c.Add Item:=cpIBM01147, Key:="IBM01147"
    c.Add Item:=cpIBM01148, Key:="IBM01148"
    c.Add Item:=cpIBM01149, Key:="IBM01149"
    c.Add Item:=cpUTF_16, Key:="utf-16"
    c.Add Item:=cpUnicodeFFFE, Key:="unicodeFFFE"
    c.Add Item:=cpWindows_1250, Key:="windows-1250"
    c.Add Item:=cpWindows_1251, Key:="windows-1251"
    c.Add Item:=cpWindows_1252, Key:="windows-1252"
    c.Add Item:=cpWindows_1253, Key:="windows-1253"
    c.Add Item:=cpWindows_1254, Key:="windows-1254"
    c.Add Item:=cpWindows_1255, Key:="windows-1255"
    c.Add Item:=cpWindows_1256, Key:="windows-1256"
    c.Add Item:=cpWindows_1257, Key:="windows-1257"
    c.Add Item:=cpWindows_1258, Key:="windows-1258"
    c.Add Item:=cpJohab, Key:="Johab"
    c.Add Item:=cpMacintosh, Key:="macintosh"
    c.Add Item:=cpX_mac_japanese, Key:="x-mac-japanese"
    c.Add Item:=cpX_mac_chinesetrad, Key:="x-mac-chinesetrad"
    c.Add Item:=cpX_mac_korean, Key:="x-mac-korean"
    c.Add Item:=cpX_mac_arabic, Key:="x-mac-arabic"
    c.Add Item:=cpX_mac_hebrew, Key:="x-mac-hebrew"
    c.Add Item:=cpX_mac_greek, Key:="x-mac-greek"
    c.Add Item:=cpX_mac_cyrillic, Key:="x-mac-cyrillic"
    c.Add Item:=cpX_mac_chinesesimp, Key:="x-mac-chinesesimp"
    c.Add Item:=cpX_mac_romanian, Key:="x-mac-romanian"
    c.Add Item:=cpX_mac_ukrainian, Key:="x-mac-ukrainian"
    c.Add Item:=cpX_mac_thai, Key:="x-mac-thai"
    c.Add Item:=cpX_mac_ce, Key:="x-mac-ce"
    c.Add Item:=cpX_mac_icelandic, Key:="x-mac-icelandic"
    c.Add Item:=cpX_mac_turkish, Key:="x-mac-turkish"
    c.Add Item:=cpX_mac_croatian, Key:="x-mac-croatian"
    c.Add Item:=cpUTF_32, Key:="utf-32"
    c.Add Item:=cpUTF_32BE, Key:="utf-32BE"
    c.Add Item:=cpX_Chinese_CNS, Key:="x-Chinese_CNS"
    c.Add Item:=cpX_cp20001, Key:="x-cp20001"
    c.Add Item:=cpX_Chinese_Eten, Key:="x_Chinese-Eten"
    c.Add Item:=cpX_cp20003, Key:="x-cp20003"
    c.Add Item:=cpX_cp20004, Key:="x-cp20004"
    c.Add Item:=cpX_cp20005, Key:="x-cp20005"
    c.Add Item:=cpX_IA5, Key:="x-IA5"
    c.Add Item:=cpX_IA5_German, Key:="x-IA5-German"
    c.Add Item:=cpX_IA5_Swedish, Key:="x-IA5-Swedish"
    c.Add Item:=cpX_IA5_Norwegian, Key:="x-IA5-Norwegian"
    c.Add Item:=cpUs_ascii, Key:="us-ascii"
    c.Add Item:=cpX_cp20261, Key:="x-cp20261"
    c.Add Item:=cpX_cp20269, Key:="x-cp20269"
    c.Add Item:=cpIBM273, Key:="IBM273"
    c.Add Item:=cpIBM277, Key:="IBM277"
    c.Add Item:=cpIBM278, Key:="IBM278"
    c.Add Item:=cpIBM280, Key:="IBM280"
    c.Add Item:=cpIBM284, Key:="IBM284"
    c.Add Item:=cpIBM285, Key:="IBM285"
    c.Add Item:=cpIBM290, Key:="IBM290"
    c.Add Item:=cpIBM297, Key:="IBM297"
    c.Add Item:=cpIBM420, Key:="IBM420"
    c.Add Item:=cpIBM423, Key:="IBM423"
    c.Add Item:=cpIBM424, Key:="IBM424"
    c.Add Item:=cpX_EBCDIC_KoreanExtended, Key:="x-EBCDIC-KoreanExtended"
    c.Add Item:=cpIBM_Thai, Key:="IBM-Thai"
    c.Add Item:=cpKoi8_r, Key:="koi8-r"
    c.Add Item:=cpIBM871, Key:="IBM871"
    c.Add Item:=cpIBM880, Key:="IBM880"
    c.Add Item:=cpIBM905, Key:="IBM905"
    c.Add Item:=cpIBM00924, Key:="IBM00924"
    c.Add Item:=cpEuc_jp, Key:="EUC-JP"
    c.Add Item:=cpX_cp20936, Key:="x-cp20936"
    c.Add Item:=cpX_cp20949, Key:="x-cp20949"
    c.Add Item:=cpCp1025, Key:="cp1025"
    c.Add Item:=cpDeprecated, Key:="deprecated"
    c.Add Item:=cpKoi8_u, Key:="koi8-u"
    c.Add Item:=cpIso_8859_1, Key:="iso-8859-1"
    c.Add Item:=cpIso_8859_2, Key:="iso-8859-2"
    c.Add Item:=cpIso_8859_3, Key:="iso-8859-3"
    c.Add Item:=cpIso_8859_4, Key:="iso-8859-4"
    c.Add Item:=cpIso_8859_5, Key:="iso-8859-5"
    c.Add Item:=cpIso_8859_6, Key:="iso-8859-6"
    c.Add Item:=cpIso_8859_7, Key:="iso-8859-7"
    c.Add Item:=cpIso_8859_8, Key:="iso-8859-8"
    c.Add Item:=cpIso_8859_9, Key:="iso-8859-9"
    c.Add Item:=cpIso_8859_13, Key:="iso-8859-13"
    c.Add Item:=cpIso_8859_15, Key:="iso-8859-15"
    c.Add Item:=cpX_Europa, Key:="x-Europa"
    c.Add Item:=cpIso_8859_8_i, Key:="iso-8859-8-i"
    c.Add Item:=cpIso_2022_jp, Key:="iso-2022-jp"
    c.Add Item:=cpCsISO2022JP, Key:="csISO2022JP"
    c.Add Item:=cpIso_2022_jp_w_1b_Kana, Key:="iso-2022-jp_w-1b-Kana"
    c.Add Item:=cpIso_2022_kr, Key:="iso-2022-kr"
    c.Add Item:=cpX_cp50227, Key:="x-cp50227"
    c.Add Item:=cpISO_2022_Trad_Chinese, Key:="ISO-2022-Traditional-Chinese"
    c.Add Item:=cpEBCDIC_Jap_Katakana_Ext, Key:="EBCDIC-Japanese-Katakana-Extended"
    c.Add Item:=cpEBCDIC_US_Can_and_Jap, Key:="EBCDIC-US-Canada-and-Japanese"
    c.Add Item:=cpEBCDIC_Kor_Ext_and_Kor, Key:="EBCDIC-Korean-Extended-and-Korean"
    c.Add Item:=cpEBCDIC_Simp_Chin_Ext, Key:="EBCDIC-Simplified-Chinese-Extended-and-Simplified-Chinese"
    c.Add Item:=cpEBCDIC_Simp_Chin, Key:="EBCDIC-Simplified-Chinese"
    c.Add Item:=cpEBCDIC_US_Can_Trad_Chin, Key:="EBCDIC-US-Canada-and-Traditional-Chinese"
    c.Add Item:=cpEBCDIC_Jap_Latin_Ext, Key:="EBCDIC-Japanese-Latin-Extended-and-Japaneseeuc_jp"
    c.Add Item:=cpEUC_CN, Key:="EUC-CN"
    c.Add Item:=cpEuc_kr, Key:="euc-kr"
    c.Add Item:=cpEUC_Traditional_Chinese, Key:="EUC-Traditional-Chinese"
    c.Add Item:=cpHz_gb_2312, Key:="hz-gb-2312"
    c.Add Item:=cpGB18030, Key:="GB18030"
    c.Add Item:=cpX_iscii_de, Key:="x-iscii-de"
    c.Add Item:=cpX_iscii_be, Key:="x-iscii-be"
    c.Add Item:=cpX_iscii_ta, Key:="x-iscii-ta"
    c.Add Item:=cpX_iscii_te, Key:="x-iscii-te"
    c.Add Item:=cpX_iscii_as, Key:="x-iscii-as"
    c.Add Item:=cpX_iscii_or, Key:="x-iscii-or"
    c.Add Item:=cpX_iscii_ka, Key:="x-iscii-ka"
    c.Add Item:=cpX_iscii_ma, Key:="x-iscii-ma"
    c.Add Item:=cpX_iscii_gu, Key:="x-iscii-gu"
    c.Add Item:=cpX_iscii_pa, Key:="x-iscii-pa"
    c.Add Item:=cpUTF_7, Key:="utf-7"
    c.Add Item:=cpUTF_8, Key:="utf-8"

    Set AllCodePages = c
End Property

Sub TestAPI()
    Dim i As Long
    Dim cpId As Variant
    Dim rndBytes As String
    rndBytes = RandomStringUnicode(1000)
    Dim convNotSupported() As Boolean
    ReDim convNotSupported(1 To 151)
    On Error Resume Next
    For Each cpId In AllCodePages
        Encode rndBytes, cpId, True
        i = i + 1
        Debug.Print i, cpId, Err.Number, Err.Description
        convNotSupported(i) = Err.Number
        On Error GoTo -1
    Next cpId
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
    s = RandomStringAlphanumeric(500000)
    Dim finds As Variant
    finds = RandomStringArray(1000, 1000, 3, 30, 255) '  '
    'finds = VBA.Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
    'finds = StringToCodepointStrings(RandomStringUnicode(6000))
    Dim replaces As Variant
    replaces = RandomStringArray(10, 10, 3, 30, 255) ' '
    'replaces = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j")
    'replaces = StringToCodepointStrings(RandomStringUnicode(5000))
    'Debug.Print ReplaceMultiple(s, finds, replaces) = ReplaceMultipleMultiPass(s, finds, replaces)
'    Debug.Print ReplaceMultiple(s, finds, replaces)
'    Debug.Print ReplaceMultipleMultiPass(s, finds, replaces)
    Dim rsingle As String
    Dim rmulti As String
    st
    rsingle = ReplaceMultiple(s, finds, replaces)
    RT "ReplaceMultiple", , True
    rmulti = ReplaceMultipleMultiPass(s, finds, replaces)
    RT "ReplaceMultipleMultiPass"
    Debug.Print rsingle = rmulti
    Debug.Print s = rsingle
    'Debug.Print ReplaceMultipleB(s, Array("1", "2", "3"), Array("44", "55"))
End Sub

Public Sub TestPrintVar()
    ' Test Case 1: Single dimensional array of integers
    Dim array1DInt(1 To 100) As Integer
    Dim i As Long
    For i = 1 To 100
        array1DInt(i) = i
    Next i
    Printf "Test Case 1: Single dimensional array of integers"
    PrintVar array1DInt
    Debug.Print vbNewLine
    
    ' Test Case 2: Single dimensional array of strings
    Dim array1DStr(1 To 3) As String
    array1DStr(1) = "Alice"
    array1DStr(2) = "Bob"
    array1DStr(3) = "Charlie"
    Debug.Print "Test Case 2: Single dimensional array of strings"
    PrintVar array1DStr
    Debug.Print vbNewLine
    
    ' Test Case 3: Two dimensional array of integers
    Const COLS_LARGE_ARRAY As Long = 30
    Const ROWS_LARGE_ARRAY As Long = 20
    Dim array2DInt() As Long
    ReDim array2DInt(1 To ROWS_LARGE_ARRAY, 1 To COLS_LARGE_ARRAY)
    Dim j As Integer
    For i = 1 To ROWS_LARGE_ARRAY
        For j = 1 To COLS_LARGE_ARRAY
            array2DInt(i, j) = i * j
        Next j
    Next i
    Debug.Print "Test Case 3: Two dimensional array of integers"
    Printf array2DInt
    Debug.Print vbNewLine
    
    Dim array2DStr() As Variant
    ReDim array2DStr(1 To ROWS_LARGE_ARRAY, 1 To COLS_LARGE_ARRAY)
    For i = 1 To ROWS_LARGE_ARRAY
        For j = 1 To COLS_LARGE_ARRAY
            array2DStr(i, j) = RandomStringAlphanumeric(Rnd * 15)
        Next j
    Next i
    Debug.Print "Test Case 4: Two dimensional array of Strings"
    PrintVar array2DStr, , , 1000
    Debug.Print vbNewLine
    
    ' Test Case 4: Two dimensional array of strings
    ReDim array2DStr(1 To 2, 1 To 2)
    array2DStr(1, 1) = "Apple"
    array2DStr(1, 2) = "Banana"
    array2DStr(2, 1) = "Cherry"
    array2DStr(2, 2) = "Durian"
    Debug.Print "Test Case 5: Two dimensional array of strings"
    PrintVar array2DStr
    Debug.Print vbNewLine
    
    ' Test Case 5: Two dimensional array of strings
    ReDim array2DStr(1 To 40, 0 To 0)
    For i = 1 To 40
        array2DStr(i, 0) = i * i
    Next i
    
    Debug.Print "Test Case 5: Two dimensional array of strings"
    PrintVar array2DStr
    Debug.Print vbNewLine
    
    ' Test Case 6: Array containing random strings with special characters
    Dim arrayRandom(1 To 3) As String
    arrayRandom(1) = RandomString(10, 33, 127) ' printable ASCII characters
    arrayRandom(2) = RandomString(10, 256, 500) ' extended ASCII characters
    arrayRandom(3) = RandomString(10, &H1F600, &H1F64F) ' emojis
    Debug.Print "Test Case 6: Array containing random strings with special characters"
    PrintVar arrayRandom, escapeNonPrintable:=False
    Debug.Print vbNewLine
    PrintVar arrayRandom, escapeNonPrintable:=True
    Debug.Print vbNewLine
    
    ' Test Case 7: Empty array
    Dim EmptyArray() As Integer
    Debug.Print "Test Case 7: Empty array"
    PrintVar EmptyArray
    Debug.Print vbNewLine
    
    Debug.Print "Test Case 8: Empty array 2"
    PrintVar Array()
    Debug.Print vbNewLine
    
    ' Test Case 8: Array containing various weird stuff
    Dim weirdArray() As Variant
    Dim nested2DimArray() As Variant
    ReDim nested2DimArray(1 To 3, 1 To 3)
    ReDim weirdArray(1 To 3, 1 To 3)
    weirdArray(1, 1) = Array(1, 2, 3, 4)
    weirdArray(1, 2) = Array(Array(1, 2, 3), 2, 3, 4)
    Set weirdArray(2, 1) = New Collection
    weirdArray(2, 2) = CCur(1000)
    weirdArray(3, 2) = nested2DimArray
    Debug.Print "Test Case 9: Weird array"
    PrintVar weirdArray
    Debug.Print vbNewLine
End Sub

Sub CompareReplaceAndReplaceMultiple()
    'In some cases ReplaceMultiple performs better than the inbuilt Replace
    'for the same task
    Const LEN_TEST_STR As Long = 1000000
    
    StartTimer
    Dim demoStr As String: demoStr = RepeatString("  a", LEN_TEST_STR / 3)
    ReadTimer "Generating test string of length " & LEN_TEST_STR, Reset:=True
    
    Dim resultNative As String: resultNative = Replace(demoStr, " ", "")
    ReadTimer "Native Replace function", Reset:=True
    
    Dim resultLib As String: resultLib = ReplaceMultiple(demoStr, " ", "")
    ReadTimer "Library ReplaceMultiple function"
    
    Debug.Print resultNative = resultLib
End Sub

Sub CompareReplaceAndReplaceB()
    Const LEN_TEST_STR As Long = 1000000
    Const REPETITIONS As Long = 1
    Dim i As Long
    
    StartTimer
    
    Dim demoStr As String: demoStr = RepeatString("  a", LEN_TEST_STR / 3)
    ReadTimer "Generating test string of length " & LEN_TEST_STR, Reset:=True
    
    For i = 1 To REPETITIONS
        Dim resultNative As String: resultNative = Replace(demoStr, " ", "  ")
    Next i
    ReadTimer "Native Replace function", Reset:=True
    
    For i = 1 To REPETITIONS
        Dim resultLib2 As String: resultLib2 = ReplaceFast(demoStr, " ", "  ")
    Next i
    ReadTimer "Library ReplaceFast function", Reset:=True
    
    For i = 1 To REPETITIONS
        Dim resultLib As String: resultLib = ReplaceB(demoStr, " ", "  ")
    Next i
    ReadTimer "Library ReplaceB function", Reset:=True
    
'    'Placeholder for a dev version of ReplaceB
'    For i = 1 To REPETITIONS
'        Dim resultLib3 As String: resultLib3 = ReplaceFastB(demoStr, " ", "  ")
'    Next i
'    ReadTimer "Library ReplaceFastB function", Reset:=True
    
    Debug.Print resultNative = resultLib2
    'Debug.Print resultLib3 = resultLib2
End Sub


Private Sub ProcessFindsUsingTrie(ByRef finds As Variant, _
                                  ByVal lCompare As VbCompareMethod)
    Dim trie As Object
    Set trie = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, j As Long, k As Long
    Dim current As Object
    Dim ch As String

    For i = LBound(finds) To UBound(finds)
        If Len(finds(i)) <> 0 Then
            'Insert current string into the trie
            Set current = trie
            For j = 1 To Len(finds(i))
                ch = Mid(finds(i), j, 1)
                If lCompare = vbTextCompare Then ch = LCase(ch)
                If Not current.Exists(ch) Then
                    current.Add ch, CreateObject("Scripting.Dictionary")
                End If
                Set current = current(ch)
            Next j
            current("end") = True
            
            'Check remaining strings against the trie
            For k = i + 1 To UBound(finds)
                If Len(finds(k)) <> 0 Then
                    Set current = trie
                    For j = 1 To Len(finds(k))
                        ch = Mid(finds(k), j, 1)
                        If lCompare = vbTextCompare Then ch = LCase(ch)
                        If Not current.Exists(ch) Then Exit For
                        Set current = current(ch)
                        
                        If current.Exists("end") And j < Len(finds(k)) Then
                            finds(k) = vbNullString
                            Exit For
                        End If
                    Next j
                End If
            Next k
        End If
    Next i
End Sub

Private Sub ProcessFindsNormally(ByRef finds As Variant, _
                                 ByVal lCompare As VbCompareMethod)
    Dim i As Long, j As Long
    For i = 0 To UBound(finds)
        If Len(finds(i)) <> 0 Then
            For j = i + 1 To UBound(finds)
                If InStr(1, finds(j), finds(i), lCompare) <> 0 Then
                    finds(j) = vbNullString
                End If
            Next j
        End If
    Next i
End Sub

Private Sub CompareFindProcessingProcedures()
    Dim tests As Variant
    Dim i As Long
    Dim original() As Variant, expected() As Variant, actual() As Variant
    
    ' Define test cases
    tests = Array( _
        Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "0"), _
        Array("hello", "helloworld", "he", "hell", "cat", "dog"), _
        Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j"), _
        Array("aa", "a", "bb", "b", "cc", "c"), _
        Array("a", "aa", "bb", "b", "c", "cc"), _
        Array("hello", "helloworld", "he", "hell", "cat", "dog"), _
        Array("hello", "ell", "orl", "World") _
    )
    
    For i = LBound(tests) To UBound(tests)
        ' Copy original array
        original = tests(i)
        
        ' Use ProcessFindsNormally as reference for the expected result
        expected = original
        st i = 0
        ProcessFindsNormally expected, vbTextCompare
        RT "ProcessFindsNormally " & i + 1
        
        ' Test ProcessFindsUsingTrie procedure
        actual = original
        st False
        ProcessFindsUsingTrie actual, vbTextCompare
        RT "ProcessFindsUsingTrie " & i + 1
        
        If Not ArraysAreEqual(actual, expected) Then
            Debug.Print "Test " & i + 1 & " failed for ProcessFindsUsingTrie procedure"
            Debug.Print "Expected: " & ToString(expected)
            Debug.Print "Actual: " & ToString(actual)
        Else
            Debug.Print "Test " & i + 1 & " passed."
        End If
    Next i
End Sub

Private Sub DemonstrateBigOcomplexityOfTrieVsNaive()
    Dim timeTaken As Currency
    st
    Dim i As Long
    Dim finds1 As Variant
    Dim finds2 As Variant
    Dim timeTrie As Currency
    Dim timeNaive As Currency
    Dim totalChars As Long
    Dim numChunks As Long
    Dim increaseFactor As Long
    increaseFactor = 2
    totalChars = 50
    numChunks = 10
    Randomize
    Do Until timeTaken > StoUs(3)
        ReDim finds1(0 To numChunks - 1)
        For i = 0 To numChunks - 1
            finds1(i) = RandomStringAlphanumeric(Rnd * totalChars \ numChunks * 2)
        Next i
        'finds1 = ChunkifyString(RandomStringAlphanumeric(totalChars), , 10)
        finds2 = finds1
        st False
        ProcessFindsUsingTrie finds1, vbTextCompare
        timeTrie = RT("Processed " & numChunks & " Finds Using Trie")
        st False
        ProcessFindsNormally finds2, vbTextCompare
        timeNaive = RT("Processed " & numChunks & _
                                      " Finds Using Nested Loop with Instr")
        
        Debug.Print "TimeTrie / TimeNaive = " & timeTrie / timeNaive
        timeTaken = timeTrie + timeNaive
        
        totalChars = totalChars * increaseFactor
        numChunks = numChunks * increaseFactor
        DoEvents
    Loop
End Sub

Private Function ArraysAreEqual(arr1 As Variant, arr2 As Variant) As Boolean
    Dim i As Long
    
    If UBound(arr1) <> UBound(arr2) Then
        ArraysAreEqual = False
        Exit Function
    End If
    
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) <> arr2(i) Then
            ArraysAreEqual = False
            Exit Function
        End If
    Next i
    
    ArraysAreEqual = True
End Function

Private Sub CompareErrorHandlingOfNativeAndApiDecoders()
    Dim s As String
    s = RandomBytes(1001)
    's = HexToString("0x6F705FEF9E1FE008BDC52A")
   ' s = HexToString("0x8B05E4950CB96D4F20F48F")
     's = HexToString("0x93191ACC480B4B614DF2FA")
    
    Dim decNative As String: decNative = DecodeUTF8(s)
    Dim decApi As String:    decApi = Decode(s, cpUTF_8)
    
    Debug.Print decNative = decApi
   ' Debug.Print StringToHex(s)
'    Debug.Print EscapeUnicode(decNative, 127)
'    Debug.Print EscapeUnicode(decApi, 127)
End Sub

'Private Sub TestErrorHandlingInTranscodingAPI()
'    Dim s As String: s = ChrW(255): s = EncodeANSI(s)
'    Dim s2 As String
'
'    Dim allCpIds As Object
'    Set allCpIds = GetAllCpIDs
'    'On Error Resume Next
'    Dim cpId As Variant
'    For Each cpId In allCpIds.Keys
'        If IsValidCodePage(CLng(cpId)) Or cpId = cpUTF_16 Then
'            s2 = Transcode(s, cpUTF_8, cpId, False, "p")
'            If s2 <> "" Then
''                s2 = Transcode(s, cpId, cpUTF_16)
''                Debug.Print "CP: " & GetAllCpIDs(cpId)
''                Debug.Print LenB(s2)
'                If Encode(GetCpInfo(cpId).defaultChar, cpId) = s2 Then
'                    Debug.Print "Used std. default char. " & GetAllCpIDs(cpId)
'                    If Not GetCpInfo(cpId).defaultChar = GetCpInfo(cpId).UnicodeDefaultChar Then
'                        Debug.Print "STD DEFAULT CHAR NOT EQUAL UNICODE DEFAULT CHAR """ & GetCpInfo(cpId).defaultChar & """ != """ & GetCpInfo(cpId).UnicodeDefaultChar & """"
'                    End If
'                ElseIf Encode(GetCpInfo(cpId).UnicodeDefaultChar, cpId) = s2 Then
'                    Debug.Print "Used unicode default char. " & GetAllCpIDs(cpId)
'                Else
'                    Debug.Print Decode(s2, cpId) & " Used OTHER character: " & StringToHex(s2) & " " & GetAllCpIDs(cpId) & "; CpName: " & GetCpInfo(cpId).CodePageName & "; Default char: """ & GetCpInfo(cpId).defaultChar & """; UnicodeDefaultChar: """ & StringToHex(GetCpInfo(cpId).UnicodeDefaultChar) & """"
'                End If
'            End If
'        Else
'            'Debug.Print "CP not supported: " & GetAllCpIDs(cpId)
'        End If
'    Next cpId
'End Sub
'

Sub TestFastReplace()
    Const NUM_LOOPS As Long = 10000
    Const LEN_STR As Long = 10
    Const lCount As Long = -1
    Dim i As Long
    Dim compareMethod As VbCompareMethod: compareMethod = vbBinaryCompare
    
    
    Dim s As String: s = RandomStringFromChars(LEN_STR, "aaa" & Space(10))
    
    StartTimer
    For i = 1 To NUM_LOOPS
        Dim resFast As String: resFast = ReplaceFast(s, "aaaaaa", "b", , lCount, compareMethod)
    Next i
    ReadTimer "ReplaceFast", , True
'    'Placeholder for a new dev version of Replace
'    For i = 1 To NUM_LOOPS
'        Dim resFaster As String: resFaster = ReplaceFaster(s, "aaaaaa", "b", , lCount, compareMethod)
'    Next i
'    ReadTimer "ReplaceFaster", , True
    For i = 1 To NUM_LOOPS
        Dim resNative As String: resNative = Replace(s, "aaaaaa", "b", , lCount, compareMethod)
    Next i
    ReadTimer "Replace Native", , True
    Debug.Print "Behavior 'ReplaceFast' is " & IIf(resFast <> resNative, _
                "not same", "same") & " as normal 'Replace'"
'    Debug.Print "Behavior 'ReplaceFaster' is " & IIf(resFaster <> resNative, _
'                "not same", "same") & " as normal 'Replace'"
    Debug.Print "Num Replacements:" & CountSubstring(s, "aaaaaa", , lCount)
'    Debug.Print resFaster
'    Debug.Print resNative
End Sub

Sub TestInStr()
    Const FIND_POS As Long = 1000000
    Const START_SEARCH_POS As Long = 1000000
    Dim str As String
    str = RandomStringFromChars(FIND_POS, "abcdefghijklmnopqrstuvwxy123") & "z" & RandomStringFromChars(FIND_POS, "abcdefghijklmnopqrstuvwxy123")
    StartTimer
    Dim posZ1 As Long: posZ1 = InStr(START_SEARCH_POS, str, "z", vbBinaryCompare)
    ReadTimer "InStr vbBinaryCompare, starting search at pos " & START_SEARCH_POS & " found at pos " & posZ1, , True
    Dim posZ2 As Long: posZ2 = InStr(START_SEARCH_POS, str, "z", vbTextCompare)
    ReadTimer "InStr vbTextCompare, starting search at pos " & START_SEARCH_POS & " found at pos " & posZ2, , True
    Dim posZ3 As Long: posZ3 = InStr(START_SEARCH_POS, LCase(Mid(str, START_SEARCH_POS - 1)), LCase("z"), vbBinaryCompare) + START_SEARCH_POS - 1
    ReadTimer "InStr vbBinaryCompare, on LCase, starting search at pos " & START_SEARCH_POS & " found at pos " & posZ3, , True
'    Dim posZ4 As Long: posZ4 = InStrTextCompare(START_SEARCH_POS, str, "z")
'    ReadTimer "InStrTextCompare, starting search at pos " & START_SEARCH_POS & " found at pos " & posZ4, , True
End Sub

Sub TestReplaceVsCountSubstring()
    Const LEN_STR As Long = 2000
    Dim s As String: s = RandomStringFromChars(LEN_STR, "Aabbbb")
    s = RepeatString("Aabbbb", LEN_STR / Len("Aabbbb"))
    StartTimer
    Dim resvbText As String: resvbText = Replace(s, "a", "b", , , vbTextCompare)
    ReadTimer "Replace Native vbTextCompare", , True
    Dim resvbBinary As String: resvbBinary = Replace(s, "a", "b", , , vbBinaryCompare)
    ReadTimer "Replace Native vbBinaryCompare", , True
    Debug.Print "Behavior 'vbText' is " & IIf(resvbText <> resvbBinary, _
                "not same", "same") & " as 'vbBinary'"
    
    Dim resInStrvbText As Long: resInStrvbText = CountSubstring(s, "a", , , vbTextCompare)
    ReadTimer "InStr vbTextCompare", , True
    Dim resInStrvbBinary As Long: resInStrvbBinary = CountSubstring(s, "a", , , vbBinaryCompare)
    ReadTimer "InStr vbBinaryCompare", , True
    Debug.Print "Behavior 'vbText' is " & IIf(resInStrvbText <> resInStrvbBinary, _
                "not same", "same") & " as 'vbBinary'"
    Debug.Print "Cound vbTextCompare: " & resInStrvbText
    Debug.Print "Cound vbBinaryCompare: " & resInStrvbBinary
End Sub

Sub TestInStrr()
    Const FIND_POS As Long = 1000000
    Const START_SEARCH_POS As Long = 1
    Dim str As String
    str = RandomStringFromChars(FIND_POS, "abcdefghijklmnopqrstuvwxy123") & "z" & RandomStringFromChars(FIND_POS, "abcdefghijklmnopqrstuvwxy123")
    
    StartTimer
    Dim posZ1 As Long: posZ1 = InStr(START_SEARCH_POS, str, "z", vbBinaryCompare)
    ReadTimer "InStr vbBinaryCompare, starting search at pos " & _
              START_SEARCH_POS & " found at pos " & posZ1, , True
    Dim posZ2 As Long: posZ2 = InStr(START_SEARCH_POS, str, "z", vbTextCompare)
    ReadTimer "InStr vbTextCompare, starting search at pos " & _
              START_SEARCH_POS & " found at pos " & posZ2, , True
End Sub

Sub TestInStrWorstCase()
    Const FIND_POS As Long = 1000000
    Const START_SEARCH_POS As Long = 1
    Dim sFind As String: sFind = "asdlkfalskdhaölskdjfölaksjfdlafsfdasdf"
    Dim str As String
    str = RepeatString(PadRight(sFind, Len(sFind) - 1), FIND_POS / Len(sFind) - 1) & _
          sFind & RepeatString(PadRight(sFind, Len(sFind) - 1), FIND_POS / Len(sFind) - 1)
    Mid$(sFind, Len(sFind) / 2, 1) = ChrW(AscW(Mid$(sFind, Len(sFind) / 2, 1)) + 1)
    str = RepeatString(sFind, FIND_POS / Len(sFind))
    Mid$(sFind, Len(sFind) / 2, 1) = ChrW(AscW(Mid$(sFind, Len(sFind) / 2, 1)) - 1)
    str = str & sFind & RepeatString(PadRight(sFind, Len(sFind) - 1), FIND_POS / Len(sFind) - 1)
    'str = RandomStringFromChars(FIND_POS, "abcdefghijklmnopqrstuvwxy123") & sFind & RandomStringFromChars(FIND_POS, "abcdefghijklmnopqrstuvwxy123")
    
    StartTimer
    Dim posZ1 As Long: posZ1 = InStr(START_SEARCH_POS, str, sFind, vbBinaryCompare)
    ReadTimer "InStr vbBinaryCompare, starting search at pos " & _
              START_SEARCH_POS & " found at pos " & posZ1, , True
    Dim posZ2 As Long: posZ2 = InStr(START_SEARCH_POS, str, sFind, vbTextCompare)
    ReadTimer "InStr vbTextCompare, starting search at pos " & _
              START_SEARCH_POS & " found at pos " & posZ2, , True
End Sub

Sub TestRnd()
    Const NUM_LOOPS As Long = 1000000
    Dim bool As Boolean
    bool = False
    Dim i As Long
    StartTimer
    For i = 1 To NUM_LOOPS
        If bool Then
            RndWH
        Else
            Rnd
        End If
    Next i
    ReadTimer
End Sub

Sub TestRndStringPerformance()
    Const NUM_LOOPS As Long = 10
    Const LEN_STR As Long = 100000

    Dim i As Long
    StartTimer
    For i = 1 To NUM_LOOPS
        RandomString LEN_STR
    Next i
    ReadTimer "RandomString with Normal Rnd", , True
    For i = 1 To NUM_LOOPS
        RandomString LEN_STR, , , True
    Next i
    ReadTimer "RandomString with RndWH", , True
    For i = 1 To NUM_LOOPS
        RandomStringAlphanumeric LEN_STR
    Next i
    ReadTimer "RandomStringAlphanumeric with Normal Rnd", , True
    For i = 1 To NUM_LOOPS
        RandomStringAlphanumeric LEN_STR, True
    Next i
    ReadTimer "RandomStringAlphanumeric with RndWH", , True
    For i = 1 To NUM_LOOPS
        RandomStringFromChars LEN_STR
    Next i
    ReadTimer "RandomStringFromChars with Normal Rnd", , True
        For i = 1 To NUM_LOOPS
        RandomStringFromChars LEN_STR, , True
    Next i
    ReadTimer "RandomStringFromChars with RndWH", , True
    For i = 1 To NUM_LOOPS
        RandomStringUnicode LEN_STR
    Next i
    ReadTimer "RandomStringUnicode with Normal Rnd", , True
    For i = 1 To NUM_LOOPS
        RandomStringUnicode LEN_STR, True
    Next i
    ReadTimer "RandomStringUnicode with RndWH", , True
    For i = 1 To NUM_LOOPS
        RandomBytes LEN_STR
    Next i
    ReadTimer "RandomBytes with Normal Rnd", , True
    For i = 1 To NUM_LOOPS
        RandomBytes LEN_STR, True
    Next i
    ReadTimer "RandomBytes with RndWH", , True
    For i = 1 To NUM_LOOPS
        RandomStringASCII LEN_STR
    Next i
    ReadTimer "RandomStringASCII with Normal Rnd", , True
    For i = 1 To NUM_LOOPS
        RandomStringASCII LEN_STR, True
    Next i
    ReadTimer "RandomStringASCII with RndWH", , True
    For i = 1 To NUM_LOOPS
        RandomStringBMP LEN_STR
    Next i
    ReadTimer "RandomStringBMP with Normal Rnd", , True
    For i = 1 To NUM_LOOPS
        RandomStringBMP LEN_STR, True
    Next i
    ReadTimer "RandomStringBMP with RndWH", , True
End Sub

Sub FindDecodeUTF8Bugs()
    Const TIMEOUT_SECONDS As Long = 10
    Const LEN_TEST_STR As Long = 3
    Const IGNORE_EDGES As Boolean = False
    
    Dim decodedBytes As Long
    Dim t As Single: t = AccurateTimerS()
    Do Until AccurateTimerS() - t > TIMEOUT_SECONDS
        DoEvents
        Dim bytes As String: bytes = RandomBytes(LEN_TEST_STR * 2, True)
        Dim apiResult As String: apiResult = Decode(bytes, cpUTF_8)
        Dim natResult As String: natResult = DecodeUTF8(bytes)
        decodedBytes = decodedBytes + LEN_TEST_STR
        
        If IGNORE_EDGES Then
            apiResult = LeftB(apiResult, LenB(apiResult) - 4)
            natResult = LeftB(natResult, LenB(apiResult))
        End If
        
        If apiResult <> natResult Then
            Debug.Print "DecodeUTF8 inconsistent with API decoder for string, "
            Debug.Print "after " & decodedBytes & " decoded bytes."
            Debug.Print "Input:         " & StringToHex(bytes)
            Debug.Print "Output API:    " & StringToHex(apiResult)
            Debug.Print "Output Native: " & StringToHex(natResult)
            Exit Sub
        End If
    Loop
End Sub


Public Function LimitConsecutiveSubstringRepetitionNaive( _
                                           ByRef str As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Dim findStr As String:    findStr = RepeatString(subStr, limit + 1)
    Dim replaceStr As String: replaceStr = RepeatString(subStr, limit)
    LimitConsecutiveSubstringRepetitionNaive = str
    If Len(findStr) = 0 Then Exit Function
    Do While InStr(1, LimitConsecutiveSubstringRepetitionNaive, _
                   findStr, Compare) > 0
        LimitConsecutiveSubstringRepetitionNaive = _
            ReplaceFast(LimitConsecutiveSubstringRepetitionNaive, findStr, _
                     replaceStr, , , Compare)
    Loop
End Function

Sub DemoLimitConsecutiveSubstringRepetition()
    'The library function is typically much faster than the naive approach and
    'has linear time complexity
    Const LEN_TEST_STR As Long = 500000
    Dim lCompare As VbCompareMethod: lCompare = vbBinaryCompare
    Dim demoStr As String
    Dim resultNaive As String
    Dim resultLib As String
    
    StartTimer
    demoStr = RandomStringFromChars(LEN_TEST_STR, "a ")
    demoStr = String(35000000, "a") & Space(10000000)
    'demoStr = RepeatString("a ", LEN_TEST_STR / 2)
    
    ReadTimer "Generating test string of length " & LEN_TEST_STR, Reset:=True
    resultNaive = LimitConsecutiveSubstringRepetitionNaive(demoStr, " ", 1, lCompare)
    ReadTimer "Naive approach", Reset:=True
    resultLib = LimitConsecutiveSubstringRepetition(demoStr, " ", 1, lCompare)
    ReadTimer "Library approach"
    
    
'    Debug.Print resultNaive = resultLib
'    Debug.Print demoStr
'    Debug.Print resultNaive
'    Debug.Print resultLib
End Sub

Public Function LimitConsecutiveSubstringRepetitionNaiveB( _
                                           ByRef str As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Dim findStr As String:    findStr = RepeatString(subStr, limit + 1)
    Dim replaceStr As String: replaceStr = RepeatString(subStr, limit)
    LimitConsecutiveSubstringRepetitionNaiveB = str
    If LenB(findStr) = 0 Then Exit Function
    Do While InStr(1, LimitConsecutiveSubstringRepetitionNaiveB, _
                   findStr, Compare) > 0
        LimitConsecutiveSubstringRepetitionNaiveB = _
            ReplaceB(LimitConsecutiveSubstringRepetitionNaiveB, findStr, _
                     replaceStr, , , Compare)
    Loop
End Function

Sub DemoLimitConsecutiveSubstringRepetitionB()
    'The library function is typically much faster than the naive approach and
    'has linear time complexity
    Const LEN_TEST_STR As Long = 1000000
    
    Dim resultNaive As String
    Dim resultLib As String
    
    StartTimer
    Dim demoStr As String: 'demoStr = RandomStringFromChars(LEN_TEST_STR, "ab ")
    'demoStr = RepeatString("  a", LEN_TEST_STR / 3)
    demoStr = String(35000000, "a") & Space(10000000)
    
    ReadTimer "Generating test string of length " & LEN_TEST_STR, Reset:=True
    resultNaive = LimitConsecutiveSubstringRepetitionNaiveB(demoStr, " ", 1)
    ReadTimer "Naive approach", Reset:=True
    resultLib = LimitConsecutiveSubstringRepetitionB(demoStr, " ", 1)
    ReadTimer "Library approach"
    Debug.Print resultNaive = resultLib
End Sub

Sub TestRandomStringFromStrings()
    Debug.Print RandomStringFromStrings(5, Array("ab", "cd"))
End Sub

Sub TestRandomStringFromCharsPerformance()
    Const STR_LENGTH As Long = 100000
    Const NUM_LOOPS As Long = 10
    Dim i As Long
    StartTimer
    For i = 1 To NUM_LOOPS
        RandomStringFromChars STR_LENGTH
    Next i
    ReadTimer "RandomStringFromChars, normal Rnd", , True
    For i = 1 To NUM_LOOPS
        RandomStringFromChars STR_LENGTH, , True
    Next i
    ReadTimer "RandomStringFromChars, custom RndWH", , True
End Sub




Sub TestIteratingString()
    Const LEN_STR As Long = 100000
    Dim lCompare As VbCompareMethod: lCompare = vbBinaryCompare
    
    Dim testStr As String
    testStr = String(LEN_STR, "a") & String(10000000, "b")
    
    Dim i As Long: i = 1
    Dim j As Long: j = 1
    StartTimer
    Do
        i = j
        j = InStr(i + 1, testStr, "a", lCompare)
    Loop Until j > i + 1 Or j = 0
    ReadTimer "InStr method", , True
    i = 0
    Do
        i = i + 1
    Loop Until StrComp(Mid$(testStr, i, 1), "a", lCompare) <> 0
    ReadTimer "StrComp method", , True
    i = 0
    Do
        i = i + 1
    Loop Until Mid$(testStr, i, 1) <> "a"
    ReadTimer "<> method", , True
End Sub

Sub TestLimitConsecutiveSubstringRepetitionPerformanceRealistic()
    Const STR_LENGTH As Long = 1000000
    Const RUN_NAIVE_AS_CHECK As Boolean = True
    
    Dim testCases As Collection
    Set testCases = New Collection
    Dim testStr As String
    
    With testCases
        testStr = String(STR_LENGTH / 2, "a") & String(STR_LENGTH / 2, "b")
        .Add VBA.Array(testStr, "a", 1)
        
        testStr = RandomStringFromChars(STR_LENGTH, "abc", True)
        .Add VBA.Array(testStr, "a", 2)
         
        testStr = RandomStringFromChars(STR_LENGTH, "abcde ", True)
        .Add VBA.Array(testStr, " ", 1)
         
        testStr = RandomStringFromStrings(STR_LENGTH, Array("asdf ", "asfdaf ", "asdfasf ", "asdfafd ", "asdfasf ", "safs ", "asdfa ", _
                "asdf ", "asfdaf ", "asdfasf ", "asdfafd ", "asdfasf ", "safs ", "asdfa ", vbCrLf, vbCrLf, vbCrLf & vbCrLf), True)
        .Add VBA.Array(testStr, vbCrLf, 1)
         
        testStr = RandomStringFromStrings(STR_LENGTH, Array("asdf ", "asfdaf ", "asdfasf ", "asdfafd ", "asdfasf ", "safs ", "asdfa ", _
                "asdf ", "asfdaf ", "asdfasf ", "asdfafd ", "asdfasf ", "safs ", "asdfa ", vbCrLf, vbCrLf, vbCrLf & vbCrLf), True)
        .Add VBA.Array(testStr, vbCrLf, 0)
    End With
    
    
    StartTimer
    Dim res1 As String, res2 As String
    Dim testParams As Variant, j As Long
    Dim totalTimeNaive As Currency
    Dim totalTimeFast As Currency
    For Each testParams In testCases
        j = j + 1
        ResetTimer
        If RUN_NAIVE_AS_CHECK Then
            res1 = LimitConsecutiveSubstringRepetitionNaive(CStr(testParams(0)), CStr(testParams(1)), testParams(2))
            totalTimeNaive = totalTimeNaive + _
                ReadTimer("Realistic case test " & j & " naive, string length " & Len(testParams(0)), , True)
        End If
        res2 = LimitConsecutiveSubstringRepetition(CStr(testParams(0)), CStr(testParams(1)), testParams(2))
        totalTimeFast = totalTimeFast + ReadTimer("Realistic case test " & j & " fast, string length " & Len(testParams(0)), , True)
        If RUN_NAIVE_AS_CHECK Then Debug.Print _
            "Test " & j & " result: " & IIf(res1 = res2, "ok", "failed!")
    Next testParams
    If RUN_NAIVE_AS_CHECK Then Debug.Print "Total time naive: " & UsToS(totalTimeNaive) & " s"
    Debug.Print "Total time fast: " & UsToS(totalTimeFast) & " s"
End Sub


Sub TestLimitConsecutiveSubstringRepetitionPerformanceWorstCase()
    Const RUN_NAIVE_AS_CHECK As Boolean = True
    Const STR_LENGTH As Long = 50000
    
    If STR_LENGTH > 50000 And RUN_NAIVE_AS_CHECK Then
        MsgBox "Dont run these tests for more than 50k characters as the naive solution has" & _
                " an O(n^2) complexity and will never return"
        Exit Sub
    End If
    
    Dim testCases As Collection
    Set testCases = New Collection
    Dim testStr As String
        
    'Add test cases:
    With testCases
        testStr = String(STR_LENGTH / 2, "a") & String(STR_LENGTH / 2, "b")
        .Add VBA.Array(testStr, "ab", 0)
        
        testStr = RepeatString("ba", STR_LENGTH / 3) & RepeatString("a", STR_LENGTH / 3)
        .Add VBA.Array(testStr, "baa", 0)
        
        Dim i As Long
        Dim insertStr As String: insertStr = RandomStringAlphanumeric(10)
        testStr = ""
        Dim insertPos As Long
        For i = 1 To STR_LENGTH / Len(insertStr)
            testStr = Insert(testStr, insertStr, insertPos)
            insertPos = RndWH * (Len(insertStr) - 2) + 1 + insertPos
        Next i
        .Add VBA.Array(testStr, insertStr, 0)
        
        testStr = ""
        Do Until Len(testStr) >= STR_LENGTH
            testStr = "a" & testStr & testStr & "b"
        Loop
        testStr = Left$(testStr, STR_LENGTH)
        .Add VBA.Array(testStr, insertStr, 0)
    End With
    
    StartTimer
    Dim res1 As String, res2 As String
    Dim testParams As Variant, j As Long
    Dim totalTimeNaive As Currency
    Dim totalTimeFast As Currency
    For Each testParams In testCases
        j = j + 1
        ResetTimer
        If RUN_NAIVE_AS_CHECK Then
            res1 = LimitConsecutiveSubstringRepetitionNaive(CStr(testParams(0)), CStr(testParams(1)), testParams(2))
            totalTimeNaive = totalTimeNaive + _
                ReadTimer("Worst case test " & j & " naive, string length " & Len(testParams(0)), , True)
        End If
        res2 = LimitConsecutiveSubstringRepetition(CStr(testParams(0)), CStr(testParams(1)), testParams(2))
        totalTimeFast = totalTimeFast + ReadTimer("Worst case test " & j & " fast, string length " & Len(testParams(0)), , True)
        If RUN_NAIVE_AS_CHECK Then Debug.Print _
            "Test " & j & " result: " & IIf(res1 = res2, "ok", "failed!")
    Next testParams
    If RUN_NAIVE_AS_CHECK Then Debug.Print "Total time naive: " & UsToS(totalTimeNaive) & " s"
    Debug.Print "Total time fast: " & UsToS(totalTimeFast) & " s"
End Sub

Sub RunLimitConsecutiveSubstringRepetitionTests()
    Dim failedTests As Long
    On Error GoTo errh:
    TestLimitConsecutiveSubstringRepetition "abababbbaabbbbb", "abb", 0
    TestLimitConsecutiveSubstringRepetition "aaaababbaababbbaaababbaababbbaaababbaababbbb", "ab", 0
    TestLimitConsecutiveSubstringRepetition "aaaabaaca", "a", 0
    TestLimitConsecutiveSubstringRepetition "baaaabaaca", "a", 0
    TestLimitConsecutiveSubstringRepetition "aaaabaaca", "a", 0
    TestLimitConsecutiveSubstringRepetition "aaaaabaaca", "a", 0
    TestLimitConsecutiveSubstringRepetition "aaaaabaaca", "aa", 0
    TestLimitConsecutiveSubstringRepetition "abaca", "aa", 0
    TestLimitConsecutiveSubstringRepetition "bababa", "ab", 0
    TestLimitConsecutiveSubstringRepetition "aaaaaaaaaaabbbbbbbbbbbbbbbbbbb", "ab", 0
    TestLimitConsecutiveSubstringRepetition "aaaabaaca", "a", 1
    TestLimitConsecutiveSubstringRepetition "baaaabaaca", "a", 1
    TestLimitConsecutiveSubstringRepetition "aaaabaaca", "a", 2
    TestLimitConsecutiveSubstringRepetition "aaaaabaaca", "a", 2
    TestLimitConsecutiveSubstringRepetition "aaaabaaca", "aa", 1
    TestLimitConsecutiveSubstringRepetition "abaca", "aa", 1
    TestLimitConsecutiveSubstringRepetition "aaaaabaaca", "aa", 1
    TestLimitConsecutiveSubstringRepetition "aaaaababaca", "ab", 1
    TestLimitConsecutiveSubstringRepetition "bbbaaababbb", "ab", 1
    TestLimitConsecutiveSubstringRepetition "bbbaaababbb", "c", 1
    TestLimitConsecutiveSubstringRepetition "erereererererr", "ere", 1
    TestLimitConsecutiveSubstringRepetition "eerr", "er", 0
    
    TestLimitConsecutiveSubstringRepetition "babaabababaababaabababaabababaabababaababaabababaaba", "babaaba", 1
    TestLimitConsecutiveSubstringRepetition "xxxxxxxxxxxxxxxxbabaabababaababaabababaabababaabababaababaabababaaba", "babaaba", 1
    TestLimitConsecutiveSubstringRepetition "babaabababaababaabababaabababaabababaababaabababaabaxxxxxxxxxxxxxxxx", "babaaba", 1
    TestLimitConsecutiveSubstringRepetition "babababababababaaaaaaaaaaaaaaaaaaaaa", "baa", 0
    TestLimitConsecutiveSubstringRepetition RepeatString("ba", 100 / 3) & "baa" & RepeatString("baaa", 100 * 2 / 3), "baa", 1
    TestLimitConsecutiveSubstringRepetition "bababababababaabaaabaaabaaabaaa" & RepeatString("baaa", 100 * 2 / 3), "baa", 1
    TestLimitConsecutiveSubstringRepetition RepeatString("babababababaaaaaaa", 1000), "baa", 0
    TestLimitConsecutiveSubstringRepetition RepeatString("ba", 100 / 3) & "a" & RepeatString("baaa", 100 * 2 / 3), _
                                            "baa", 0
    TestLimitConsecutiveSubstringRepetition "abababababababWl3aWl3aTWl3Wl3aWl3aTIceGHTIceGHaTIceGHIceGHTIceGH", "Wl3aTIceGH", 0
    TestLimitConsecutiveSubstringRepetition "Wl3aTIcWl3aWl3aTIceGWl3aTIcWl3aTIceGHeGHHTIceGHeGHbabababababababa", "Wl3aTIceGH", 0
    TestLimitConsecutiveSubstringRepetition "abababababababWl3aTIcWl3aWl3aTIceGWl3aTIcWl3aTIceGHeGHHTIceGHeGHbabababababababa", "Wl3aTIceGH", 0
    TestLimitConsecutiveSubstringRepetition "abaabaababaabaababaabaababaabaababaabaab", "aba", 1
    TestLimitConsecutiveSubstringRepetition "aabbabababbabababbabababbabababbababbababbababbababbababbababbababbababbababbabababbabababbab", "babab", 1
    TestLimitConsecutiveSubstringRepetition "bbabbbbabbabbbbabbabbbbabbabbbbabba", "bbabb", 1
    TestLimitConsecutiveSubstringRepetition "bbaaaababab", "aab", 0
    
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
        Debug.Print "LimitConsecutiveSubstringRepetition FAILED " & failedTests & " tests!"
    End If
    Exit Sub
errh:
    If Err.Number = vbObjectError + 43233 Then
        failedTests = failedTests + 1
        Debug.Print Err.Description
        Resume Next
    Else
        Err.Raise Err
    End If
End Sub

Sub StressTestLimitConsecutiveSubstringRepetition()
    Const TIMEOUT_SECONDS As Long = 5000
    Const MAX_REPETITIONS As Long = 10
    Const TEST_STR_LENGTH As Long = 20
    Const NUM_SUBSTRINGS As Long = 4
    Const MAX_LEN_SUBSTRING As Long = 10
    Const NUM_DIFFERENT_CHARACTERS As Long = 2
    Const PRINT_EVERY_N_SECONDS As Long = 3
    
    Dim totalTime As Currency: totalTime = AccurateTimerS
    Dim totalTimeNaive As Currency, totalTimeFast As Currency
    Dim i As Long
    Dim numTestsTotal As Long
    Dim printTimer As Currency
    Do Until AccurateTimerS - totalTime > TIMEOUT_SECONDS
        DoEvents
        Dim sourceArr As Variant
        sourceArr = RandomStringArray(NUM_SUBSTRINGS, MAX_LEN_SUBSTRING, 3, 97, 97 + NUM_DIFFERENT_CHARACTERS - 1, True)
        Dim maxRepetitions As Long: maxRepetitions = MAX_REPETITIONS + 1
        Dim subStrToLimit As String
        subStrToLimit = RepeatString(CStr(sourceArr(Int(RndWH * NUM_SUBSTRINGS))), _
                                     Int(RndWH * 2) + 1)
        For i = LBound(sourceArr) To UBound(sourceArr)
            sourceArr(i) = RepeatString(CStr(sourceArr(i)), 1 + Int(RndWH * (maxRepetitions - 1)))
        Next i
        Dim testStr As String
        testStr = RandomStringFromStrings(TEST_STR_LENGTH, sourceArr, True)
        Dim maxRepsLimit As Long: maxRepsLimit = Int(RndWH * 3) + Int(RndWH * 3)
        Dim startTimeNaive As Currency: startTimeNaive = AccurateTimerUs
        Dim controlStr As String
        controlStr = LimitConsecutiveSubstringRepetitionNaive(testStr, subStrToLimit, maxRepsLimit, vbBinaryCompare)
        totalTimeNaive = totalTimeNaive + AccurateTimerUs - startTimeNaive
        Dim startTimeFast As Currency: startTimeFast = AccurateTimerUs
        Dim resStr As String
        resStr = LimitConsecutiveSubstringRepetition(testStr, subStrToLimit, maxRepsLimit, vbBinaryCompare)
        totalTimeFast = totalTimeFast + AccurateTimerUs - startTimeFast
        numTestsTotal = numTestsTotal + 1
        If AccurateTimerS - printTimer > PRINT_EVERY_N_SECONDS Then
            printTimer = AccurateTimerS
            Debug.Print numTestsTotal & " tests performed so far. Total runtime " & _
                        AccurateTimerS - totalTime & " s"
            Debug.Print "Time spent naive: " & UsToS(totalTimeNaive), "Time spent fast: " & UsToS(totalTimeFast)
        End If
        'If Len(controlStr) < Len(testStr) Then Debug.Print "String got shorter" Else Debug.Print "Not shorter"
        If controlStr <> resStr Then
            Debug.Print "Bug found!"
            If Len(testStr) <= 1000 Then Debug.Print "testStr: " & testStr
            Debug.Print "subStrToLimit: " & subStrToLimit
            Debug.Print "limit: " & maxRepsLimit
            Stop
            Exit Sub
        End If
    Loop
    Debug.Print "Time spent naive: " & UsToS(totalTimeNaive)
    Debug.Print "Time spent fast: " & UsToS(totalTimeFast)
End Sub

