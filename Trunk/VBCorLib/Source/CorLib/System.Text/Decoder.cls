VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Decoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: Decoder
'

''
' Provides an interface to be used in converting blocks of bytes into unicode characters.
'
Option Explicit

''
' Returns the number of characters a block of bytes can produce.
'
' @param Bytes The bytes to be used in decoding to characters.
' @param Index The starting index in <i>Bytes</i> to begin decoding from.
' @param Count The number of bytes to be used in the decoding process.
' @return The number of characters that can be produced from <i>Count</i> number of bytes.
'
Public Function GetCharCount(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, Optional ByVal Flush As Boolean) As Long: End Function

''
' Returns the decoded unicode characters from a block of bytes.
'
' @param Bytes The bytes to be used in decoding to characters.
' @param ByteIndex The starting index to begin decoding bytes into characters.
' @param ByteCount The number of bytes to be used in the decoding process.
' @param Chars The output array for the decoded unicode characters.
' @param CharIndex The starting index in the array to begin storing characters.
' @return The number of characters actually decoded.
'
Public Function GetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long: End Function

''
' Returns a string representation of this instance.
'
' @return A string that describes this instance.
'
Public Function ToString() As String: End Function

''
' Returns an indication of equality with the passed in value.
'
' @param value The value to test equality with.
' @param The result of the equality test.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number that describes this instance.
'
' @return A pseudo-unique number that describes this instance.
'
Public Function GetHashCode() As Long: End Function

