Attribute VB_Name = "TestLibStringTools"
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
        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef Frequency As Currency) As LongPtr
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

Private Function GetTickCount() As Currency
    #If Mac Then
        GetTickCount = mach_continuous_time()
    #Else
        QueryPerformanceCounter GetTickCount
    #End If
End Function

Private Function GetFrequency() As Currency
    #If Mac Then
        Dim timebaseInfo As MachTimebaseInfo
        mach_timebase_info timebaseInfo
        GetFrequency = (timebaseInfo.Denominator / timebaseInfo.Numerator) * 100000#
    #Else
        QueryPerformanceFrequency GetFrequency
    #End If
End Function

Private Function AccurateTimer() As Currency
    AccurateTimer = GetTickCount / GetFrequency
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
End Sub

Private Sub TestEncodersAndDecoders()
    Const STR_LENGTH As Long = 1000000
    
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



