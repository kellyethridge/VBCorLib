VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringComparerStatic"
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
' Module: StringComparerStatic
'

''
' Provides a set of <a href="StringComparer.html">StringComparer</a> objects implementing different string comparison rules.
'
' @see StringComparer
'
Option Explicit

Private mOrdinal            As New OrdinalStringComparer
Private mOrdinalIgnoreCase  As New OrdinalStringComparer


''
' Gets a <a href="StringComparer.html">StringComparer</a> object that provides case-sensitive string comparisons.
'
' @return A <a href="StringComparer.html">StringComparer</a> object that is case-sensitive.
' @remarks The comparison is done in a similar fashion as using the standard equality check of two VisualBasic strings in a variant culture manner.
' <p>The <a href="StringComparer.html">StringComparer</a> object can be retrieved using the following code.</p>
' <pre>
'   Dim c As StringComparer
'   Set c = StringComparer.Ordinal
' </pre>
' <p>The instance of the <a href="StringComparer.html">StringComparer</a> object returned is a private class and simply implements the <a href="StringComparer.html">StringComparer</a> interface.</p>
Public Property Get Ordinal() As StringComparer
    Set Ordinal = mOrdinal
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that provides case-insensitive string comparisons.
'
' @return A <a href="StringComparer.html">StringComparer</a> that is case-insensitive.
' @remarks The comparison is done like using the <b>StrComp</b> function to check two VisualBasic strings in a variant culture manner.
' <p>The <a href="StringComparer.html">StringComparer</a> object can be retrieved using the following code.</p>
' <pre>
'   Dim c As StringComparer
'   Set c = StringComparer.OrdinalIgnoreCase
' </pre>
' <p>The instance of the <a href="StringComparer.html">StringComparer</a> object returned is a private class and simply implements the <a href="StringComparer.html">StringComparer</a> interface.</p>
Public Property Get OrdinalIgnoreCase() As StringComparer
    Set OrdinalIgnoreCase = mOrdinalIgnoreCase
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mOrdinal.Init False
    mOrdinalIgnoreCase.Init True
End Sub
