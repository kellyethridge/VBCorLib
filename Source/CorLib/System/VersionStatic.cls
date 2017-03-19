VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VersionStatic"
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
' Module: VersionStatic
'

''
' Provides static methods to help create Version objects.
'
' @see Version
'
Option Explicit

Private mParser As New VersionParser

''
' Parses a string in the format of major.minor[.build[.revision]].
'
' @param s The string to be parsed.
' @return A version defined from the string.
' @remarks Atleast 2 of the version numbers must be supplied, and
' all version numbers must be non-negative.
'
Public Function Parse(ByRef s As String) As Version
    If Not mParser.TryParse(s, Parse) Then
        mParser.ThrowError
    End If
End Function

''
' Tries to parse a string in the format of major.minor[.build[.revision]] returning a boolean value
' indicating the success or failure of parsing. The parsed version is returned
' through the <b>ByRef Result</b> parameter supplied.
'
' @param s The string containing the version information to be parsed.
' @param Result The <b>ByRef</b> parameter used to pass the <pre>Version</pre>
' back with on a successful parse.
' @return <b>True</b> if the <i>input</i> was successfully parsed; otherwise <b>False</b>..
'
Public Function TryParse(ByRef s As String, ByRef Result As Version) As Boolean
    TryParse = mParser.TryParse(s, Result)
End Function
