Attribute VB_Name = "LibStringToolsTestFunctions"
'===============================================================================
' VBA StringTools - Test Functions
' ------------------------------------------------------------------------------------
' https://github.com/guwidoe/VBA-StringTools/blob/main/src/test/LibStringToolsTestFunctions.bas
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

'This module contains functions that are required for the LibStringTools Tests
'module. These are functions may be used in test procedures for various reasons.
'For example, some algorithms implemented in LibStringTools have a trivial
'second implementation here that is much simpler but also much less efficient.
'These simple algorithms are very useful to test the output of the library
'functions for correctness, even if they are not efficient enough for other
'usage.

Option Explicit


'Function to compare the output of LimitConsecutiveSubstringRepetition
'against. The algorithm of this function is trivially correct but can be very
'inefficient in some cases.
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

'Function to compare the output of LimitConsecutiveSubstringRepetitionB
'against. The algorithm of this function is trivially correct but can be very
'inefficient in some cases.
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

'Alternative for LimitConsecutiveSubstringRepetitionNaive
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

Public Function ArraysAreEqual(ByRef arr1 As Variant, _
                                ByRef arr2 As Variant, _
                       Optional ByVal requireSameIndexBase As Boolean = True) _
                                As Boolean
    ArraysAreEqual = False
    If requireSameIndexBase Then
        If LBound(arr1) <> LBound(arr2) Or UBound(arr1) <> UBound(arr2) Then
            Exit Function
        End If
    Else
        If UBound(arr1) - LBound(arr1) <> UBound(arr2) - LBound(arr2) Then
            Exit Function
        End If
    End If
    
    Dim i As Long
    Dim j As Long: j = LBound(arr2)
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) <> arr2(j) Then Exit Function
        j = j + 1
    Next i
    ArraysAreEqual = True
End Function
