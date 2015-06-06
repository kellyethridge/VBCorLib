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
' This interface is used to allow for custom enumerators to be built.
'
' @remarks
' <p>The enumerator can be used manually to enumerate through the
' values of a list, or it can be wrapped in a custom enumerator to allow it to be
' used in the For..Each context.</p>
' <p>To use an IEnumerator object in For..Each, call the CreateEnumerator method
' passing in the IEnumerator object. The returned value is then returned in
' the standard NewEnum(-4) function called by For..Each. The enumeration will then
' be delegated to the custom IEnumerator object.<br><br>
' <b>Example</b><br>
' <pre>
' Public Function NewEnum() As IUnknown
'     Dim e As New MyEnumerator
'     e.Init Me
'     Set NewEnum = CreateEnumerator(e)
' End Function<br>
' </pre>
'
Option Explicit

''
' This returns the current item of the enumeration. The current
' item is called after a successful MoveNext.
'
' @return The current value in the enumeration.
'
Public Property Get Current() As Variant: End Property

''
' Requests the enumerator moves to the next item being enumerated.
' Returns True on success, False otherwise. This is called prior
' to Current. If this is successful, Current will be called.
'
' @return Indication if the move was successful.
'
Public Function MoveNext() As Boolean: End Function

''
' Requests the enumerator resets itself to begin enumerating from the beginning.
'
Public Sub Reset(): End Sub

