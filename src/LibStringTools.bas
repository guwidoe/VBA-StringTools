Attribute VB_Name = "LibStringTools"
Option Explicit

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
'Make HexToString and ReplaceUnicodeLiterals Mac compatible by removing Regex

#If Mac = 0 Then
    #If VBA7 Then
        Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    #Else
        Private Declare Function MultiByteToWideChar Lib "kernel32" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
        Private Declare Function WideCharToMultiByte Lib "kernel32" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    #End If
#End If

#If Mac = 0 Then
'Returns strings defined as hex literal as string.
'Accepts the following formattings:
'0xXXXXXX (even number of Xes, X = 0-9 or a-f, not case sensitive)
'XXXXXX (even number of Xes, X = 0-9 or a-f, not case sensitive)
'X XX XX X XX (Xes separated by " ", X = 0-9 or a-f, not case sensitive)
'instead of " ", the following separators are also accepted: ",;-+"
'Accepts any combination of the above formattings
'e.g.: "0x610062006300" will be converted to "abc"
Public Function HexToString(ByVal hexStr As String) As String
    Const methodName As String = "HexToString"
    Dim s As String
    
    s = " " & Replace(Replace(Replace(Replace(Replace(LCase(hexStr), _
            "0x", " "), ",", " "), ";", " "), "-", " "), "+", " ") & " "
            
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = False 'Already LCase()
        .Pattern = " ([a-f0-9]) "
        s = .Replace(s, "0$1 ")
    End With
    
    s = Replace(s, " ", "")
    If Len(s) Mod 2 Then Err.Raise 5, methodName, _
            "Invalid Hex string literal. (Length is not even)"
    
    Dim mask As String: mask = Replace(Space(Len(s)), " ", "[a-f0-9]")
    
    If Not s Like mask Then Err.Raise 5, methodName, _
        "Invalid Hex string literal. (Contains characters other than a-f & 0-9)"
    
    Dim i As Long
    Dim b() As Byte: ReDim b(0 To Len(s) \ 2 - 1)
    
    For i = LBound(b) To UBound(b)
        b(i) = "&H" & Mid$(s, i * 2 + 1, 2)
    Next i
    HexToString = b
End Function
#End If

'Converts the input string into a string of hex literals.
'e.g.: "abc" will be turned into "0x610062006300" (UTF-16LE)
Public Function StringToHex(ByVal str As String) As String
    Dim i As Long
    Dim b() As Byte:      b = str
    Dim hexStr As String: hexStr = "0x" & Space(Len(str) * 4 + 2)

    For i = 1 To UBound(b) + 1
        Mid(hexStr, i * 2 + 1, 2) = Right$("0" & Hex$(b(i - 1)), 2)
    Next i
    StringToHex = hexStr
End Function

#If Mac = 0 Then
'Replaces all occurences of unicode literals of the following formattings:
'\uXXXX \UXXXX (4 or 8 hex digits, 8 for chars outside BMP) (X = 0-9 or a-f)
'u+XXXX U+XXXX (4 or 5 hex digits) (X = 0-9 or a-f)
'&#dddd;       (1 to 6 dec digits) (d = 0-9)
'e.g.: the string "abc &#97 u+0062 \U0063" will be transformed to "abc a b c"
'This function can be slow for very large amount of different literals and very
'long input strings
'Depends on: ChrU
Public Function ReplaceUnicodeLiterals(ByVal str As String) As String
    Const PATTERN_UNICODE_LITERALS As String = _
        "\\u000[0-9a-f]{5}|\\u[0-9a-f]{4}|u\+[0-9|a-f]{4,5}|&#\d{1,6}"
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
        If Left(mv, 1) = "&" Then
            codepoint = CLng(Mid(mv, 3, Len(mv) - 3))
        Else
            codepoint = CLng("&H" & Mid(mv, 3))
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
'with literals of the following formatting:
'\uXXXX (4 or 8 hex digits, 8 for chars outside BMP) (X = 0-9 or a-f)
'e.g.: No example possible because VBE doesn't allow such characters
'Depends on: AscU
Public Function EncodeUnicodeCharacters(ByVal str As String) As String
    Dim codepoint As Long
    Dim i As Long
    Dim j As Long:          j = 1
    Dim result() As String: ReDim result(1 To Len(str))
    
    
    For i = 1 To Len(str)
        codepoint = AscW(Mid(str, i, 1)) And &HFFFF&
        
        If codepoint >= &HD800& Then codepoint = AscU(Mid(str, i, 2))
        
        If codepoint > &HFFFF& Then 'Outside BMP
            result(j) = "\u" & "000" & Hex(codepoint)
            i = i + 1
        ElseIf codepoint > &HFF Then 'BMP
            result(j) = "\u" & Right("00" & Hex(codepoint), 4)
        Else
            result(j) = Mid(str, i, 1)
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
Public Function AscU(ByVal char As String) As Long
    Dim s As String
    Dim lo As Long
    Dim hi As Long
    
    If Len(char) = 1 Then
        AscU = AscW(char) And &HFFFF&
    Else
        s = Left(char, 2)
        hi = AscW(Mid(s, 1, 1)) And &HFFFF&
        lo = AscW(Mid(s, 2, 1)) And &HFFFF&
        
        If &HDC00& > lo Or lo > &HDFFF& Then
            AscU = hi
            Exit Function
        End If
        AscU = (hi - &HD800&) * &H400& + (lo - &HDC00&) + &H10000
    End If
End Function

'Function transcoding an ANSI encoded string to the VBA-native UTF-16LE
Public Function DecodeANSI(ByVal ansiStr As String) As String
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
Public Function EncodeANSI(ByVal utf16leStr As String) As String
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
            ansi(i) = &H3F
            j = j + 2 '&H3F = "?"
        End If
    Next i
    EncodeANSI = ansi
End Function

'Slower but shorter version
Public Function EncodeANSI_2(ByVal utf16leStr As String) As String
    Dim i As Long
    Dim ansi() As Byte: ReDim ansi(1 To Len(utf16leStr))
    
    For i = 1 To UBound(ansi)
        ansi(i) = Asc(Mid(utf16leStr, i, 1))
    Next i
    EncodeANSI_2 = ansi
End Function

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-8
Public Function EncodeUTF8(ByVal utf16leStr As String, _
                  Optional ByVal raiseErrors As Boolean = True) As String
    Const methodName As String = "EncodeUTF8"
    Dim codepoint As Long
    Dim lowSurrogate As Long
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0
    Dim utf8() As Byte:       ReDim utf8(Len(utf16leStr) * 4 - 1)
    
    Do While i <= Len(utf16leStr)
        codepoint = AscW(Mid(utf16leStr, i, 1)) And &HFFFF&
        
        If codepoint >= &HD800& And codepoint <= &HDBFF& Then 'high surrogate
            lowSurrogate = AscW(Mid(utf16leStr, i + 1, 1)) And &HFFFF&
            
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
    EncodeUTF8 = MidB$(utf8, 1, j)
End Function

'Function transcoding an UTF-8 encoded string to the VBA-native UTF-16LE
Public Function DecodeUTF8(ByVal utf8Str As String, _
                  Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "DecodeUTF8"
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
                Dim m As Long:            m = codepoint - &H10000
                Dim loSurrogate As Long:  loSurrogate = &HDC00& Or (m And &H3FF)
                Dim hiSurrogate As Long:  hiSurrogate = &HD800& Or (m \ &H400&)
                
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
    DecodeUTF8 = MidB$(utf16, 1, j)
End Function

#If Mac = 0 Then
'Transcoding a VBA-native UTF-16LE encoded string to UTF-8 using ADODB.Stream
'Much faster than EncodeUTF8, but only available on Windows
Public Function EncodeUTF8_2(ByVal utf16leStr As String) As String
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText utf16leStr
        .Position = 0
        .Type = 1 ' adTypeBinary
        .Position = 3 ' Skip BOM (Byte Order Mark)
        EncodeUTF8_2 = .Read
        .Close
    End With
End Function

'Transcoding an UTF-8 encoded string to VBA-native UTF-16LE using ADODB.Stream
'Faster than EncodeUTF8 for medium length strings but only available on Windows
'Warning: This function performs extremely slow for strings bigger than ~5MB
Public Function DecodeUTF8_2(ByVal utf8Str As String) As String
    Dim b() As Byte: b = utf8Str
    With CreateObject("ADODB.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .Write b
        .Position = 0
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        DecodeUTF8_2 = .ReadText
        .Close
    End With
End Function

'Transcoding a VBA-native UTF-16LE encoded string to UTF-8 using the Windows API
'Much faster than EncodeUTF8 and faster than EncodeUTF8_2, (Windows only)
Public Function EncodeUTF8_3(ByVal utf16leStr As String) As String
    Const CP_UTF8 As Long = 65001
    Dim utf8Len As Long
    Dim utf8() As Byte

    utf8Len = _
        WideCharToMultiByte(CP_UTF8, 0, StrPtr(utf16leStr), -1, 0, 0, 0, 0) - 1
    If utf8Len <= 0 Then Exit Function
    ReDim utf8(utf8Len - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(utf16leStr), -1, VarPtr(utf8(0)), _
                                                                   utf8Len, 0, 0
    EncodeUTF8_3 = utf8
End Function

'Transcoding an UTF-8 encoded string to VBA-native UTF-16LE using the Windows API
'Much faster than DecodeUTF8 and faster than DecodeUTF8_2, (Windows only)
Public Function DecodeUTF8_3(ByVal utf8Str As String) As String
    Const CP_UTF8 As Long = 65001
    Dim sLen As Long
    Dim utf8() As Byte: utf8 = utf8Str
    
    sLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LenB(utf8Str), 0, 0)
    If sLen <= 0 Then Exit Function
    
    DecodeUTF8_3 = String$(sLen, 0)
    MultiByteToWideChar CP_UTF8, 0, VarPtr(utf8(0)), LenB(utf8Str), _
                                                      StrPtr(DecodeUTF8_3), sLen
End Function
#End If

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-32
Public Function EncodeUTF32LE(ByVal utf16leStr As String, _
                     Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "EncodeUTF32LE"
    
    If utf16leStr = "" Then Exit Function
    
    Dim codepoint As Long
    Dim lowSurrogate As Long
    Dim utf32() As Byte:      ReDim utf32(Len(utf16leStr) * 4 - 1)
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0
    
    Do While i <= Len(utf16leStr)
        codepoint = AscW(Mid(utf16leStr, i, 1)) And &HFFFF&
        
        If codepoint >= &HD800& And codepoint <= &HDBFF& Then 'high surrogate
            lowSurrogate = AscW(Mid(utf16leStr, i + 1, 1)) And &HFFFF&
            
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
Public Function DecodeUTF32LE(ByVal utf32str As String, _
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
    ReDim Preserve utf16(LBound(utf16) To j - 1)
    DecodeUTF32LE = utf16
End Function

'Function returning a string containing all alphanumeric characters equally
'distributed. (0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ)
Public Function RandomStringAlphanumeric(ByVal Length As Long) As String
    If Length < 1 Then Exit Function
    
    Dim i As Long
    Dim char As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)
    
    Randomize
    For i = 0 To Length - 1
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
Public Function RandomStringAlphanumeric2(ByVal Length As Long) As String
    Dim a As Variant
    If IsEmpty(a) Then
        a = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", _
                  "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", _
                  "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", _
                  "w", "x", "y", "z", _
                  "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    End If
    
    Dim i As Long
    Dim result As String: result = Space(Length)
    
    Randomize
    For i = 1 To Length
        Mid(result, i, 1) = a(Int(Rnd() * 62))
    Next i
    RandomStringAlphanumeric2 = result
End Function

'Function returning a string containing all characters from the BMP
'(Basic Multilingual Plane, all 2 byte UTF-16 chars) equally, randomly
'distributed. Excludes surrogate range and BOM.
Public Function RandomStringBMP(ByVal Length As Long) As String
    Const MAX_UINT As Long = &HFFFF&
    
    If Length < 1 Then Exit Function
    
    Dim i As Long
    Dim char As Long
    Dim b() As Byte:  ReDim b(0 To Length * 2 - 1)
    
    Randomize
    For i = 0 To Length - 1
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

'Function returning a string containing all valid unicode characters equally,
'randomly distributed. Excludes surrogate range and BOM.
Public Function RandomStringUnicode(ByVal Length As Long) As String
    'Length in UTF-16 codepoints, not unicode codepoints!
    Const MAX_UNICODE As Long = &H10FFFF
    
    If Length < 1 Then Exit Function
    
    Dim s As String
    Dim i As Long
    
    Dim char As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)
    
    Randomize
    If Length > 1 Then
        For i = 0 To Length - 2
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
        Mid(s, Len(s), 1) = ChrW(Int(Rnd() * &HFFFE& + 1))
    RandomStringUnicode = s
End Function

'Function returning a string containing all ASCII characters equally,
'randomly distributed.
Public Function RandomStringASCII(Length As Long) As String
    Const MAX_ASC As Long = &H7F&
    Dim i As Long
    Dim char As Integer
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)
    
    Randomize
    For i = 0 To Length - 1
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
        sChr = Mid(str, i, 1)
        
        If InStr(1, inklChars, sChr, vbBinaryCompare) Then
            Mid(str, j, 1) = sChr
            j = j + 1
        End If
    Next i
    CleanString = Left(str, j - 1)
End Function

#If Mac = 0 Then
'Removes all non-numeric characters from a string.
'Only keeps codepoints U+0030 - U+0039
Public Function RegExNumOnly(s As String) As String
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
        sChr = Mid(str, i, 1)
        If sChr Like "#" Then _
            Mid(str, j, 1) = sChr: j = j + 1
    Next i
    RemoveNonNumeric = Left(str, j - 1)
End Function

'Inserts a string into another string at a specified position
'Insert("abcd", "ff", 0) = "ffabcd"
'Insert("abcd", "ff", 1) = "affbcd"
'Insert("abcd", "ff", 2) = "abffcd"
'Insert("abcd", "ff", 3) = "abcffd"
'Insert("abcd", "ff", 4) = "abcdff"
'Insert("abcd", "ff", 9) = "abcdff"
Public Function Insert(str As String, _
                       strToInsert As String, _
                       afterPos As Long) As String
    If afterPos < 0 Then afterPos = 0
    
    Insert = Mid(str, 1, afterPos) & strToInsert & Mid(str, afterPos + 1)
End Function

'Splits a string at every occurrence of the specified delimiter "delim", unless
'that delimiter occurs between non-escaped quotes. e.g. (" asf delim asdf ")
'will not be split. Quotes will not be removed.
'Quotes can be excaped by repetition.
'E.g.: Splits string:
'                      "Hello "" ""World" "Goodbye World"
'               into:
'                      "Hello "" "" World"
'               and:
'                      "Goodbye World"
Public Function SplitUnlessInQuotes(ByVal str As String, _
                           Optional ByVal delim As String = " ", _
                           Optional limit As Long = -1) As Variant
    Dim i As Long
    Dim s As String
    Dim ub As Long:         ub = -1
    Dim parts As Variant:   ReDim parts(0 To 0)
    Dim doSplit As Boolean: doSplit = True
    
    For i = 1 To Len(str)
        If ub = limit - 2 Then
            ub = ub + 1
            ReDim Preserve parts(0 To ub)
            parts(ub) = Mid(str, i)
            Exit For
        End If
        
        If Mid(str, i, 1) = """" Then doSplit = Not doSplit
        
        If Mid(str, i, Len(delim)) = delim And doSplit Or i = Len(str) Then
            If i = Len(str) Then s = s & Mid(str, i, 1)
            
            ub = ub + 1
            ReDim Preserve parts(0 To ub)
            parts(ub) = s
            s = ""
            i = i + Len(delim) - 1
        Else
            s = s & Mid(str, i, 1)
        End If
    Next i
    SplitUnlessInQuotes = parts
End Function

'Adds fillerChars to the right side of a string to make it the specified length
Public Function ReDimPreserveString(str As String, _
                              ByVal Length As Long, _
                     Optional ByVal fillerChar As String = " ") As String
    If Length > Len(str) Then
        ReDimPreserveString = str & String(Length - Len(str), fillerChar)
    Else
        ReDimPreserveString = Left(str, Length)
    End If
End Function
