VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderFallbackBuffer"
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
' Module: EncoderFallbackBuffer
'

''
' Provides a buffer that allows a fallback handler to return an alternate string to an encoder when it cannot encode an input character.
'
'@Folder("CorLib.System.Text")
Option Explicit

''
' When implemented in a derived class, gets the number of characters in the current EncoderFallbackBuffer object that remain to be processed.
'
' @return The number of characters in the current fallback buffer that have not yet been processed.
' @remarks The Fallback method returns <b>True</b> if <b>Remaining</b> is a nonzero value.
' For an exception fallback handler, the Remaining method should return zero.
Public Property Get Remaining() As Long: End Property

''
' When overridden in a derived class, prepares the fallback buffer to handle the specified surrogate pair.
'
' @param UnknownCharOrHighSurrogate Can be a non-surrogate charactor or a high surrogate of a surrogate pair.
' @param UnknownLowSurrogate The low surrogate matched with the high surrogate of the surrogate pair.
' @param Index The index position of the surrogate pair in the input buffer.
' @return <b>True</b> if the fallback buffer can process <i>UnknownCharOrHighSurrogate</i> as a non-surrogate or
' process <i>UnknownCharOrHighSurrogate</i> and <i>LowSurrogate</i> as a surrogate pair; <b>False</b> if the fallback buffer ignores the characters.
Public Function Fallback(ByVal UnknownCharOrHighSurrogate As Long, Optional ByRef UnknownLowSurrogate As Variant, Optional ByVal Index As Long) As Boolean: End Function

''
' When implemented in a derived class, retrieves the next character in the fallback buffer.
'
' @return The next character in the fallback buffer.
'
Public Function GetNextChar() As Integer: End Function

''
' When implemented in a derived class, causes the next call to the GetNextChar method to access the data buffer character position that is prior to the current character position.
'
' @return <b>True</b> if the MovePrevious operation was successful; otherwise, <b>False</b>.
'
Public Function MovePrevious() As Boolean: End Function

''
' Initializes all data and state information pertaining to this fallback buffer.
'
' @remarks After calling the Reset method, the Remaining property should return 0, the GetNextChar method should return U+0000, and the MovePrevious method should return <b>False</b>.
'
Public Sub Reset(): End Sub

''
' Determines whether the specified value passed in is equal to the current object.
'
' @param Value The value to compare to the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
' @see IObject
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string that represents the current object.
'
' @return A string the represents the current object.
' @see IObject
'
Public Function ToString() As String: End Function
