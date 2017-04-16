VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextReader"
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
' Module: TextReader
'

''
' Provides an interface to read text characters from a stream.
'
' @see TextReaderStatic
' @see StreamReader
' @see StringReader
'
Option Explicit

''
' Closes the current reader.
'
Public Sub CloseReader(): End Sub

''
' Returns if the value is equal to this instance of <b>TextReader</b>
'
' @param Value The value to campare equality with.
' @return <b>True</b> if <i>Value</i> equals this class instance, else <b>False</b>.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number that identifies this object instance.
'
' @return Value that identifies this instance.
'
Public Function GetHashCode() As Long: End Function

''
' Returns the next character from the stream without advancing to the next characters.
'
' @return The next character in the Stream, or -1 if no characters remain in the stream.
'
Public Function Peek() As Long: End Function

''
' Returns the next character in the stream, consuming that character.
'
' @return The next character in the stream, or -1 if no more characters are in the stream.
'
Public Function Read() As Long: End Function

''
' Fills a character array (Integer Array) with the specified number of characters, returning
' the actual number of characters read from the stream.
'
' @param Chars The character array to receive <i>Count</i> characters from the stream.
' @param Index The starting index in <i>Chars</i> to begin storing characters.
' @param Count The number of characters to be read from the stream.
' @return The actual number of characters read from the stream.
' @remarks Characters are encoded within the stream. A single character may require more that
' one byte to represent it. This may result in fewer characters read than there are bytes in the stream.
'
Public Function ReadBlock(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long: End Function

''
' Reads characters from the stream until a Return or a LineFeed is found, returning all
' characters up to that line break.
'
' @return A string containing characters up to the line break.
' @remarks The line break is not included in the return value. If a Return is followed
' immediately by a LineFeed, then the LineFeed is also removed.
' <p>If the return value is a null-string, then the end of the stream was reached. To
' determine if the return string is a null-string us the <b>cString.IsNull</b> function.</p>
'
Public Function ReadLine() As String: End Function

''
' Reads the remaining characters in the stream, returning them as a string.
'
' @return A string containing the remaining characters in the stream.

Public Function ReadToEnd() As String: End Function

''
' Returns a string representation of this instance.
'
' @return A string representation of this instance.
'
Public Function ToString() As String: End Function

