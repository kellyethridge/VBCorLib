VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderFallbackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: DecoderFallbackBuffer
'

''
' Provides a buffer that allows a fallback handler to return an alternate string to a decoder when it cannot decode an input byte sequence.
'
' @remarks An encoding defines a mapping between a Unicode character and an encoded sequence of bytes. A decoding operation, which converts an input byte sequence to an output character, fails if no mapping is defined for a particular byte sequence.
' <p>VBCorLib provides a failure handling mechanism, called a fallback, if a conversion cannot be performed. All decoder fallback handlers must implement the following:
' <ul>
'   <li>A decoder fallback, which is represented by a class that implements the <a href="DecoderFallback.html">DecoderFallback</a> interface.</li>
'   <li>A decoder fallback buffer, which is represented by a type that implements the <b>DecoderFallbackBuffer</b> interface that can return a string to the conversion operation.</li>
' </ul></p>
Option Explicit

''
' When implemented by a class, prepares the fallback buffer to handle the specified input byte sequence.
'
' @param BytesUnknown An input array of bytes.
' @param Index The index position of a byte in <i>BytesUnknown</i>.
' @return <b>True</b> if the fallback buffer can process <i>BytesUnknown</i>; <b>False</b> if the fallback buffer ignores <i>BytesUnknown</i>.
Public Function Fallback(ByRef BytesUnknown() As Byte, ByVal Index As Long) As Boolean: End Function

''
' When implemented by a class, retrieves the next character in the fallback buffer.
'
' @return The next character in the fallback buffer.
Public Function GetNextChar() As Integer: End Function

''
' When implemented by a class, causes the next call to the <a href="DecoderFallbackBuffer.GetNextChar.html">GetNextChar</a> method
' to access the data buffer character position that is prior to the current character position.
'
' @return <b>True</b> if the <b>MovePrevious</b> operation was successful; otherwise, <b>False</b>.
Public Function MovePrevious() As Boolean: End Function

''
' When implemented by a class, gets the number of characters in the current <a href="DecoderFallbackBuffer.html">DecoderFallbackBuffer</a> object that remain to be processed.
'
' @return The number of characters in the current fallback buffer that have not yet been processed.
Public Property Get Remaining() As Long: End Property

''
' Initializes all data and state information pertaining to this fallback buffer.
'
' @remarks <b>Notes to Implementers</b>
' <p>After calling the <b>Reset</b> method, the <a href="DecoderFallbackBuffer.Get.Remaining.html">Remaining</a> property should return 0,
' the <a href="DecoderFallbackBuffer.GetNextChar.html">GetNextChar</a> method should return U+0000,
' and the <a href="DecoderFallbackBuffer.MovePrevious.html">MovePrevious</a> method should return <b>False</b>.</p>
Public Sub Reset(): End Sub

''
' Determines whether the specified value passed in is equal to the current object.
'
' @param Value The value to compare to the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
' @see IObject
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
' @see IObject
Public Function GetHashCode() As Long: End Function

''
' Returns a string that represents the current object.
'
' @return A string the represents the current object.
' @see IObject
Public Function ToString() As String: End Function
