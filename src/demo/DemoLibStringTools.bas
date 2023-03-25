Attribute VB_Name = "DemoLibStringTools"
Option Explicit

'''=============================================================================
''' VBA StringTools Demo Module
''' ------------------------------------------
''' https://github.com/guwidoe/VBA-StringTools
''' ------------------------------------------
''' MIT License
'''
''' Copyright (c) 2023 Guido Witt-Döring
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

'--------------------------------------------------------------------|
#If VBA7 Then                                                       '|
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As LongPtr                                                    '|
                                                                    '|
    Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef counter As Currency) As LongPtr                                                     '|  THESE API FUNCTIONS ARE
#Else                                                               '|  FOR THE ACCURATE TIMER
    Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long                                                       '|
                                                                    '|
    Private Declare Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long                                                       '|
#End If                                                             '|
'--------------------------------------------------------------------|

Sub DemonstrateHexString()
    Dim utf16leTestHexString As String
    utf16leTestHexString = "0x3DD800DE3DD869DC0D203DD869DC3ED8B2DD3DD869DC3DD869DC0D203DD869DC0D203DD867DC0D203DD866DC3ED8B2DD0D203DD869DC0D203DD867DC0D203DD866DC3ED8B2DD0D203DD867DC0D203DD866DC55006E00690063006F006400650053007500700070006F007200740000D800DC6500730074003DD800DE0D203DD869DC3DD869DC0D203DD869DC0D203DD867DC0D203DD866DC3DD881DC3CD8FCDF0D2040260FFE3ED8D4DD3CD8FBDF0D2042260FFE3DD869DC0D2064270FFE0D203DD868DC3CD8C3DF3CD8FBDF0D2040260FFE"
    
    Dim s As String
    s = HexToString(utf16leTestHexString)

    'Write the string full of emojis to the worksheet "Sheet1"
    ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = s

    'Convert the UTF16 hex representation to UTF-8:
    s = EncodeUTF8(HexToString(utf16leTestHexString))

    'Look at the UTF8 bytes in the immediate window
    s = StringToHex(s)
    Debug.Print s

    'Convert UTF-8 hex string to regular vba string (UTF-16LE)
    s = DecodeUTF8_3(HexToString(s))
    
    ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = s

    'Confirm it is still the same as before:
    Debug.Assert s = ThisWorkbook.Worksheets("Sheet1").Cells(1, 1)

    'Convert all characters outside the ANSI range to Unicode literals:
    s = EncodeUnicodeCharacters(s)
    
    'Print the encoded string
    Debug.Print s
    
    'Convert back and check if it stayed the same
    s = ReplaceUnicodeLiterals(s)
    Debug.Assert s = ThisWorkbook.Worksheets("Sheet1").Cells(1, 1)
End Sub

Sub TestEncodersAndDecoders()
    Const STR_LENGTH As Long = 1000000
    Dim fullUnicode As String
    Dim bmpUnicode As String '(Basic Multilingual Plane)
    Dim utf16AsciiOnly As String
    fullUnicode = RandomStringUnicode(STR_LENGTH)
    bmpUnicode = RandomStringBMP(STR_LENGTH)
    utf16AsciiOnly = RandomStringASCII(STR_LENGTH)
    
    'VBA natively implemented Encoders/Decoders
    Debug.Print "UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF8(EncodeUTF8(bmpUnicode)) = bmpUnicode, "passed", "failed")
        
    #If Mac = 0 Then
    Debug.Print "UTF-8 Encoder/Decoder 2 Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF8_2(EncodeUTF8_2(bmpUnicode)) = bmpUnicode, "passed", "failed")
        
    Debug.Print "UTF-8 Encoder/Decoder 3 Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF8_3(EncodeUTF8_2(bmpUnicode)) = bmpUnicode, "passed", "failed")
    #End If
    
    Debug.Print "UTF-32 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF32LE(EncodeUTF32LE(bmpUnicode)) = bmpUnicode, "passed", "failed")
        
    Debug.Print "UTF-8 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF8(EncodeUTF8(fullUnicode)) = fullUnicode, "passed", "failed")
    
    #If Mac = 0 Then
    Debug.Print "UTF-8 Encoder/Decoder 2 Test full Unicode: " & _
        IIf(DecodeUTF8_2(EncodeUTF8_2(fullUnicode)) = fullUnicode, "passed", "failed")
        
    Debug.Print "UTF-8 Encoder/Decoder 3 Test full Unicode: " & _
        IIf(DecodeUTF8_3(EncodeUTF8_3(fullUnicode)) = fullUnicode, "passed", "failed")
    #End If
    
    Debug.Print "UTF-32 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF32LE(EncodeUTF32LE(fullUnicode)) = fullUnicode, "passed", "failed")
        
    Debug.Print "ANSI Encoder/Decoder Test: " & _
        IIf(DecodeANSI(EncodeANSI(utf16AsciiOnly)) = utf16AsciiOnly, "passed", "failed")
End Sub

Sub TestUTF8EncodersPerformance()
    Dim startTime As Currency, endTime As Currency
    Dim perSecond As Currency, timeElapsed As Double
    getFrequency perSecond
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant, strLengths As Variant
    numRepetitions = VBA.Array(100000, 1000, 10)
    strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String, s As String
    Dim numReps As Long, strLength As Long, i As Long, j As Long
    For i = LBound(numRepetitions) To UBound(numRepetitions)
        numReps = numRepetitions(i)
        strLength = strLengths(i)
    
        s = RandomStringUnicode(strLength)
        's = RandomStringBMP(strLength)
        's = RandomStringASCII(strLength)
        
        description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-8 Encoder:
        getTime startTime
        For j = 1 To numReps
            EncodeUTF8 s
        Next j
        getTime endTime
        timeElapsed = (endTime - startTime) / perSecond
        Debug.Print "EncodeUTF8 took: " & timeElapsed & description
        
        #If Mac = 0 Then
            'ADODB.Stream UTF-8 Encoder:
            getTime startTime
            For j = 1 To numReps
                EncodeUTF8_2 s
            Next j
            getTime endTime
            timeElapsed = (endTime - startTime) / perSecond
            Debug.Print "EncodeUTF8_2 took: " & timeElapsed & description
            
            'Windows API UTF-8 Encoder:
            getTime startTime
            For j = 1 To numReps
                EncodeUTF8_3 s
            Next j
            getTime endTime
            timeElapsed = (endTime - startTime) / perSecond
            Debug.Print "EncodeUTF8_3 took: " & timeElapsed & description
        #End If
        DoEvents
    Next i
End Sub


Sub TestUTF8DecodersPerformance()
    Dim startTime As Currency, endTime As Currency
    Dim perSecond As Currency, timeElapsed As Double
    getFrequency perSecond
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant, strLengths As Variant
    numRepetitions = VBA.Array(100000, 1000, 10)
    strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String, s As String
    Dim numReps As Long, strLength As Long, i As Long, j As Long
    For i = LBound(numRepetitions) To UBound(numRepetitions)
        numReps = numRepetitions(i)
        strLength = strLengths(i)
    
        s = RandomStringUnicode(strLength)
        's = RandomStringBMP(strLength)
        's = RandomStringASCII(strLength)
        
        s = EncodeUTF8(s)
        description = " seconds to encode a string of length " & _
                      strLength & " " & numReps & " times."
                      
        'VBA Native UTF-8 Decoder:
        getTime startTime
        For j = 1 To numReps
            DecodeUTF8 s
        Next j
        getTime endTime
        timeElapsed = (endTime - startTime) / perSecond
        Debug.Print "DecodeUTF8 took: " & timeElapsed & description
        
        #If Mac = 0 Then
            'ADODB.Stream UTF-8 Decoder:
            getTime startTime
            For j = 1 To numReps
                DecodeUTF8_2 s
            Next j
            getTime endTime
            timeElapsed = (endTime - startTime) / perSecond
            Debug.Print "DecodeUTF8_2 took: " & timeElapsed & description
            
            'Windows API UTF-8 Decoder:
            getTime startTime
            For j = 1 To numReps
                DecodeUTF8_3 s
            Next j
            getTime endTime
            timeElapsed = (endTime - startTime) / perSecond
            Debug.Print "DecodeUTF8_3 took: " & timeElapsed & description
        #End If
        DoEvents
    Next i
End Sub

Sub TestUTF32EncodersAndDecodersPerformance()
    Dim startTime As Currency, endTime As Currency
    Dim perSecond As Currency, timeElapsed As Double
    getFrequency perSecond
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant, strLengths As Variant
    numRepetitions = VBA.Array(100000, 1000, 10)
    strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String, s As String, s2 As String
    Dim numReps As Long, strLength As Long, i As Long, j As Long
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
        getTime startTime
        For j = 1 To numReps
            EncodeUTF32LE s
        Next j
        getTime endTime
        timeElapsed = (endTime - startTime) / perSecond
        Debug.Print "EncodeUTF32LE took: " & timeElapsed & description
        

        'VBA Native UTF-32 Decoder:
        getTime startTime
        For j = 1 To numReps
            DecodeUTF32LE s2
        Next j
        getTime endTime
        timeElapsed = (endTime - startTime) / perSecond
        Debug.Print "DecodeUTF32LE took: " & timeElapsed & description

        DoEvents
    Next i
End Sub

Sub TestANSIEncodersAndDecodersPerformance()
    Dim startTime As Currency, endTime As Currency
    Dim perSecond As Currency, timeElapsed As Double
    getFrequency perSecond
    Application.EnableCancelKey = xlInterrupt
    
    Dim numRepetitions As Variant, strLengths As Variant
    numRepetitions = VBA.Array(100000, 1000, 10)
    strLengths = VBA.Array(100, 1000, 1000000)
    
    Dim description As String, s As String, s2 As String
    Dim numReps As Long, strLength As Long, i As Long, j As Long
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
        getTime startTime
        For j = 1 To numReps
            EncodeANSI s
        Next j
        getTime endTime
        timeElapsed = (endTime - startTime) / perSecond
        Debug.Print "EncodeANSI took: " & timeElapsed & description
        

        'VBA Native UTF-32 Decoder:
        getTime startTime
        For j = 1 To numReps
            DecodeANSI s2
        Next j
        getTime endTime
        timeElapsed = (endTime - startTime) / perSecond
        Debug.Print "DecodeANSI took: " & timeElapsed & description

        DoEvents
    Next i
End Sub

Sub TestDifferentWaysOfGettingNumericalValuesFromStrings()
    Dim t As Single
    Dim str As String
    t = Timer()
    
    str = RandomStringAlphanumeric(5000000)
    'str = RandomStringAlphanumeric2(5000000)
    
    Debug.Print "Creating string took " & Timer - t & " seconds"
    
    t = Timer()
    Debug.Print Len(RemoveNonNumeric(str))
    Debug.Print "RemoveNonNumeric took " & Timer - t & " seconds"

    t = Timer()
    Debug.Print Len(CleanString(str, "0123456789"))
    Debug.Print "CleanString took " & Timer - t & " seconds"
    
    t = Timer()
    Debug.Print Len(RegExNumOnly(str))
    Debug.Print "RegExNumOnly took " & Timer - t & " seconds"
End Sub


