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
    s = DecodeUTF8(HexToString(s))
    
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


