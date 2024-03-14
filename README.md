# VBA-StringTools

Useful methods for interaction with strings in VBA, including transcoding, converting, escaping, and many highly performant utility functions. The library is designed to be fully cross-platform and cross-application, meaning it should run in any VBA environment.

## Installation

Just import the provided code module into your project:

- [LibStringTools.bas](https://github.com/guwidoe/VBA-StringTools/blob/main/src/LibStringTools.bas)

## Usage

A demo module is in the making, but for now, all exposed methods are preceded by a banner comment explaining the functionality and serving as documentation.

List of public/exposed methods:

- ArrayReplaceMultiple
- AscU
- ChrU
- ChunkifyString
- CleanString
- ColLetterToNumber
- CountSubstring
- CountSubstringB
- CountSubstringUnlessEscaped
- CountSubstringUnlessEscapedB
- Decode
- DecodeANSI
- DecodeASCII
- DecodeUTF32LE
- DecodeUTF8
- Encode
- EncodeANSI
- EncodeASCII
- EncodeUTF32LE
- EncodeUTF8
- EscapeUnicode
- GetBstrFromWideStringPtr (Windows only)
- GetNonUnicodeSystemCodepage
- HexToString
- Insert
- InsertB
- IntegersToString
- LimitConsecutiveSubstringRepetition
- LimitConsecutiveSubstringRepetitionB
- PadLeft
- PadLeftB
- PadRight
- PadRightB
- ParseDate
- Printf
- RandomBytes
- RandomString
- RandomStringAlphanumeric
- RandomStringArray
- RandomStringASCII
- RandomStringBMP
- RandomStringFromChars
- RandomStringFromStrings
- RandomStringUnicode
- RemoveNonNumeric
- RepeatString
- ReplaceB
- ReplaceFast
- ReplaceMultiple
- ReplaceMultipleB
- ReplaceMultipleMultiPass
- SetPrintfSettings
- SplitB
- SplitUnlessEscaped
- SplitUnlessEscapedB
- SplitUnlessInQuotes
- StringToCodepointNums
- StringToCodepointStrings
- StringToHex
- StringToIntegers
- ToString
- Transcode
- TrimX
- UnescapeUnicode

## Notes

- No extra library references are needed (e.g. Microsoft Scripting Runtime)
- Works in any host Application (Excel, Word, AutoCAD, etc.)
- Works on both Windows and Mac. Only the function `GetBstrFromWideStringPtr` is currently limited to Windows.
- Works in both x32 and x64 application environments

## License

MIT License

Copyright (c) 2023 Guido Witt-Dörring

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
