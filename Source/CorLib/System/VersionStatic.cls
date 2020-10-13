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
' Provides static methods to help create <a href="Version.html">Version</a> objects.
'
' @see Version
'
Option Explicit

Private mParser As New VersionParser

''
' Converts the string representation of a version number to an equivalent <a href="VersionStatic.html">VersionStatic</a> object.
'
' @param s A string that contains a version number to convert.
' @return An object that is equivalent to the version number specified in the <i>s</i> parameter.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>s</i> has fewer than two or more than four version components.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td>At least one component in <i>s</i> is less than zero.</td>
'    </tr>
'    <tr>
'        <td><a href="FormatException.html">FormatException</a></td><td>At least one component in <i>s</i> is not an integer.</td>
'    </tr>
'    <tr>
'        <td><a href="OverflowException.html">OverflowException</a></td><td>At least one component in <i>s</i> represents a number that is greater than the max 32-bit value.</td>
'    </tr>
' </table>
' @remarks The <i>s</i> parameter must have the following format:
' <pre>
'   Major.Minor[.Build[.Revision]]
' </pre>
' where <i>Major</i>, <i>Minor</i>, <i>Build</i>, and <i>Revision</i> are the string representations of the version number's four components: major version number, minor version number, build number, and revision number.
' Optional components are shown in square brackets ([ and ]). The components must appear in the specified order, and must be separated by periods.
Public Function Parse(ByRef s As String) As Version
    If Not mParser.TryParse(s, Parse) Then
        mParser.ThrowError
    End If
End Function

''
' Tries to convert the string representation of a version number to an equivalent <a href="Version.html">Version</a> object, and returns a value that indicates whether the conversion succeeded.
'
' @param s A string that contains a version number to convert.
' @param Result When this method returns, contains the <a href="Version.html">Version</a> equivalent of the number that is contained in <i>s</i>,
' if the conversion succeeded, or <b>Nothing</b> if the conversion failed.
' @return <b>True</b> if the <i>s</i> parameter was converted successfully; otherwise <b>False</b>..
Public Function TryParse(ByRef s As String, ByRef Result As Version) As Boolean
    TryParse = mParser.TryParse(s, Result)
End Function
