VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEnumerator"
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
' Module: IEnumerator
'

''
' Supports a simple iteration over a collection.
'
' @remarks The enumerator can be used manually to enumerate through the values of a list, or it
' can be wrapped in a custom enumerator to allow it to be used in the <i>For..Each</i> context.
' <p>To use an <b>IEnumerator</b> object in For..Each, call the <a href="Constructors.CreateEnumerator.html">CreateEnumerator</a> method
' passing in the <b>IEnumerator</b> object. The returned value is then returned in the standard NewEnum(Procedure ID: -4) function called
' by <i>For..Each</i>. The enumeration will then be delegated to the custom <b>IEnumerator</b> object.</p>
' @see IEnumerable
' @include "..\DocIncludes\IEnumerator.html"
Option Explicit

''
' Gets the current element in the collection.
'
' @return The current element in the collection.
' @remarks After an enumerator is created or after the <a href="IEnumerator.Reset.html">Reset</a> method is called, the <a href="IEnumerator.MoveNext.html">MoveNext</a>
' method must be called to advance the enumerator to the first element of the collection before reading the value of the <b>Current</b> property; otherwise, <b>Current</b> is undefined.
' <p><b>Current</b> also throws an exception if the last call to <a href="IEnumerator.MoveNext.html">MoveNext</a> returned <b>False</a>, which indicates the end of the collection.</p>
' <p><b>Current</b> does not move the position of the enumerator, and consecutive calls to <b>Current</b> return the same object until either
' <a href="IEnumerator.MoveNext.html">MoveNext</a> or <a href="IEnumerator.Reset.html">Reset</a> is called.</p>
' @include "..\DocIncludes\IEnumerator.html"
Public Property Get Current() As Variant: End Property

''
' Advances the enumerator to the next element of the collection.
'
' @return <b>True</b> if the enumerator was successfully advanced to the next element; <b>False</b> if the enumerator has passed the end of the collection.
' @remarks After an enumerator is created or after the <a href="IEnumerator.Reset.html">Reset</a> method is called, an enumerator is positioned before the
' first element of the collection, and the first call to the <b>MoveNext</b> method moves the enumerator over the first element of the collection.
' <p>If <b>MoveNext</b> passes the end of the collection, the enumerator is positioned after the last element in the collection and <b>MoveNext</b> returns <b>False</b>.
' When the enumerator is at this position, subsequent calls to <b>MoveNext</b> also return <b>False</b> until <a href="IEnumerator.Reset.html">Reset</a> is called.</p>
' @include "..\DocIncludes\IEnumerator.html"
Public Function MoveNext() As Boolean: End Function

''
' Requests the enumerator resets itself to begin enumerating from the beginning.
'
' @include "..\DocIncludes\IEnumerator.html"
Public Sub Reset(): End Sub

