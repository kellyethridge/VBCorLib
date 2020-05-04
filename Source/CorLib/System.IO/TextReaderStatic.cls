VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextReaderStatic"
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
' Module: TextReaderStatic
'

''
' Provides static methods and properties for the <a href="TextReader.html">TextReader</a> class.
'
' @see TextReader
Option Explicit

Private mNullReader As New NullTextReader

''
' Returns a <a href="TextReader.html">TextReader</a> with no data to read from.
'
' @return The <b>Null</b> text reader.
' @remarks Reader from the <b>NullReader</b> text reader is similar to reading from the end of a stream:
' <ul>
'   <li>The <a href="TextReader.Read.html">Read</a> and <a href="TextReader.Peek.html">Peek</a> methods return -1.</li>
'   <li>The <a href="TextReader.ReadBlock.html">ReadBlock</a> method returns zero.</li>
'   <li>The <a href="TextReader.ReadLine.html">ReadLine</a> method returns <b>vbNullString</b>.</li>
'   <li>The <a href="TextReader.ReadToEnd.html">ReadToEnd</a> method returns <b>vbNullString</b>.</li>
' </ul>
' <p>Access this property using <code>TextReader.NullReader</code>.</p>
Public Property Get NullReader() As TextReader
    Set NullReader = mNullReader
End Property

