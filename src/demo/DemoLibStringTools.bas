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


Sub TestEncodersAndDecoders()
    Const STR_LENGTH As Long = 5000000
    Dim fullUnicode As String
    Dim bmpUnicode As String '(Basic Multilingual Plane)
    Dim utf16AsciiOnly As String
    fullUnicode = RandomStringUnicode(STR_LENGTH)
    bmpUnicode = RandomStringBMP(STR_LENGTH)
    utf16AsciiOnly = RandomStringASCII(STR_LENGTH)
    
    'VBA natively implemented Encoders/Decoders
    Debug.Print "UTF-8 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF8(EncodeUTF8(bmpUnicode)) = bmpUnicode, "passed", "failed")
    Debug.Print "UTF-32 Encoder/Decoder Test Basic Multilingual Plane: " & _
        IIf(DecodeUTF32(EncodeUTF32(bmpUnicode)) = bmpUnicode, "passed", "failed")
        
    Debug.Print "UTF-8 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF8(EncodeUTF8(fullUnicode)) = fullUnicode, "passed", "failed")
    Debug.Print "UTF-32 Encoder/Decoder Test full Unicode: " & _
        IIf(DecodeUTF32(EncodeUTF32(fullUnicode)) = fullUnicode, "passed", "failed")
        
    Debug.Print "ANSI Encoder/Decoder Test: " & _
        IIf(DecodeANSI(EncodeANSI(utf16AsciiOnly)) = utf16AsciiOnly, "passed", "failed")
End Sub

Sub TestEncodersAndDecodersPerformance()
    Dim startTime As Currency
    Dim endTime As Currency
    Dim perSecond As Currency
    Dim timeElapsed As Double
                                
    getFrequency perSecond
    
    Const NUM_REPETITIONS_SHORT_STRING As Long = 1000000
    Const STR_LENGTH_SHORT_STRING As Long = 100
    
    Const NUM_REPETITIONS_MEDIUM_STRING As Long = 1000
    Const STR_LENGTH_MEDIUM_STRING As Long = 10000
    
    Const NUM_REPETITIONS_LONG_STRING As Long = 1
    Const STR_LENGTH_SHORT_STRING As Long = 10000000
    
    Dim fullUnicode As String
    Dim bmpUnicode As String '(Basic Multilingual Plane)
    Dim utf16AsciiOnly As String
    fullUnicode = RandomStringUnicode(STR_LENGTH)
    bmpUnicode = RandomStringBMP(STR_LENGTH)
    utf16AsciiOnly = RandomStringASCII(STR_LENGTH)
    
    'UTF-8 Encoder:
    getTime startTime
    For i = 1 To NUM_REPETITIONS
        s2 = DecodeUTF8(s)
    Next i
    getTime endTime
    timeElapsed = (endTime - startTime) / perSecond
    Debug.Print "Code 1 took: " & timeElapsed & " Seconds"

    getTime startTime
    For i = 1 To NUM_REPETITIONS
        Debug.Print s2 = DecodeUTF82(s)
    Next i
    getTime endTime
    timeElapsed = (endTime - startTime) / perSecond
    Debug.Print "Code 2 took: " & timeElapsed & " Seconds"
    
    getTime startTime
    For i = 1 To NUM_REPETITIONS
        s = ChrU(i, True)
    Next i
    getTime endTime
    timeElapsed = (endTime - startTime) / perSecond
    Debug.Print "Code 3 took: " & timeElapsed & " Seconds"
    Debug.Print " "
End Sub



Sub test()
    Dim testFilePath As String
    testFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator
    Dim testFileFullName As String
    testFileFullName = testFilePath & "UnicodeTest.txt"
    
    Dim utf16leTestString As String
    utf16leTestString = "0x3DD800DE3DD869DC0D203DD869DC3ED8B2DD3DD869DC3DD869DC0D203DD869DC0D203DD867DC0D203DD866DC3ED8B2DD0D203DD869DC0D203DD867DC0D203DD866DC3ED8B2DD0D203DD867DC0D203DD866DC55006E00690063006F006400650053007500700070006F007200740000D800DC6500730074003DD800DE0D203DD869DC3DD869DC0D203DD869DC0D203DD867DC0D203DD866DC3DD881DC3CD8FCDF0D2040260FFE3ED8D4DD3CD8FBDF0D2042260FFE3DD869DC0D2064270FFE0D203DD868DC3CD8C3DF3CD8FBDF0D2040260FFE"
    
    Dim s As String
    s = Utf16LeHexToString(utf16leTestString)
    Dim s2 As String
    s2 = s
    s = RandomStringFullUnicode(1000000)
    Debug.Print DecodeUTF32(EncodeUTF32(s)) = s
    
    'PutBytes(testFileFullName) = EncodeUTF82(ChrU(&HFFFD))
    
 'ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = Utf16LeHexToString(testStr)
 
End Sub

Sub TestDifferentWaysOfGettingNumericalValuesFromStrings()
    Dim t As Single
    Dim str As String
    t = Timer()
    str = RandomStringAlphanumeric2(5000000)
    'str = RandomStringAlphanumeric(5000000)
    
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


