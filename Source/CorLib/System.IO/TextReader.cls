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
'@Folder("CorLib.System.IO")
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
' @return The next character in the stream, or -1 if no characters remain in the stream.
' @remarks The <b>Peek</b> method returns an integer value in order to determine whether the end of the file, or another error has occurred.
' This allows a user to first check if the returned value is -1 before casting it to a <b>Integer</b> type.
' <p>The current position of the <a href="TextReader.html">TextReader</a> is not changed by this operation. The returned value is -1 if no more characters are available.</p>
' <p>If the returned value is not -1 then use the <a href="Char.ConvertFromInt32.html">Char.ConvertFromInt32</a> method to convert to an <b>Integer</b> character.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The <a href="TextReader.html">TextReader</a> is closed.</td>
'   </tr>
' </table>
Public Function Peek() As Long: End Function

''
' Reads the next character from the input stream and advances the character position by one character.
'
' @return The next character from the input stream, or -1 if no more characters are available.
' @remarks If the returned value is not -1 then use the <a href="Char.ConvertFromInt32.html">Char.ConvertFromInt32</a> method to convert to an <b>Integer</b> character.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The <a href="TextReader.html">TextReader</a> is closed.</td>
'   </tr>
' </table>
Public Function Read() As Long: End Function

''
' Reads a maximum of <i>Count</i> characters from the current stream, and writes the data to <i>Chars</i>, beginning at <i>Index</i>.
'
' @param Chars The character array to receive <i>Count</i> characters from the stream.
' @param Index The starting index in <i>Chars</i> to begin storing characters.
' @param Count The maximum number of characters to be read from the stream.
' @return The number of characters that have been read. The number will be less than or equal to <i>Count</i>, depending on whether all input characters have been read.
' @remarks Characters are encoded within the stream. A single character may require more that
' one byte to represent it. This may result in fewer characters read than there are bytes in the stream.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Chars</i> is uninitialized.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>The <i>Char</i> length minus <i>Index</i> is less than <i>Count</i>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Index</i> or <i>Count</i> is negative.</td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The <a href="TextReader.html">TextReader</a> is closed.</td>
'    </tr>
' </table>
Public Function ReadBlock(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long: End Function

''
' Reads a line of characters from the current stream and returns the data as a string.
'
' @return The next line from the input stream, or <b>vbNullString</b> if all characters have been read.
' @remarks A line is defined as a sequence of characters followed by a carriage return (0x000d), a line feed (0x000a), a carriage return followed by
' a line feed, <a href="Environment.Get.NewLine.html">Environment.NewLine</a>, or the end of stream marker. The string that is returned does not
' contain the terminating carriage return and/or line feed. The returned value is <b>vbNullString</b> if the end of the input stream has been reached.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The <a href="TextReader.html">TextReader</a> is closed.</td>
'   </tr>
' </table>
Public Function ReadLine() As String: End Function

''
' Reads all characters from the current position to the end of the <a href="TextReader.html">TextReader</a> and returns them as one string.
'
' @return A string containing all characters from the current position to the end of the <a href="TextReader.html">TextReader</a>.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The <a href="TextReader.html">TextReader</a> is closed.</td>
'   </tr>
' </table>
Public Function ReadToEnd() As String: End Function

''
' Returns a string representation of this instance.
'
' @return A string representation of this instance.
'
Public Function ToString() As String: End Function

