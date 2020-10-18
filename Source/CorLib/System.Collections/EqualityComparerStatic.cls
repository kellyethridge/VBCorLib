VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EqualityComparerStatic"
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
' Module: EqualityComparerStatic
'

''
' Provides static methods to support <a href="IEqualityComparer.html">IEqualityComparer</a> interface.
'
Option Explicit

Private mDefault As New EqualityComparer

''
' Returns a default equality comparer.
'
' @return A default equality comparer.
' @remarks <p>The default <a href="IEqualityComparer.html">IEqualityComparer</a> implementation
' simply forwards the <a href="IEqualityComparer.Equals.html">Equals</a> and <a href="IEqualityComparer.GetHashCode.html">GetHashCode</a>
' methods to the <a href="ObjectStatic.Equals.html">Object.Equals</a> and <a href="ObjectStatic.GetHashCode.html">Object.GetHashCode</a> methods.</p>
' <p>Access this property using <b>EqualityComparer.Default</b>.
Public Property Get Default() As IEqualityComparer
    Set Default = mDefault
End Property
