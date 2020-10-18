VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICollection"
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
' Module: ICollection
'

''
' Provides an interface to access size and enumeration capabilities of collections.
'
' @remarks<br>
' <p>Custom collection classes should implement this interface to allow
' VBCorLib to interact with it in way necessary. Method calls that add
' a range will usually accept objects implementing this interface.</p>
' <p>Classes implementing this interface should also implement the
' IEnumerable interface to maintain consistency with inherited interfaces.
' And, the author should consider implementing the NewEnum and GetEnumerator
' functions in the classes public interface.</p>
'
' @see IEnumerator
'
Option Explicit

''
' Returns an IEnumerator object.
'
' @return The enumerator.
'
Public Function GetEnumerator() As IEnumerator: End Function

''
' Returns an IEnumVariant object.
'
' @return The enumerator
'
Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

''
' Returns the size of the collection.
'
' @return The size of the collection.
'
Public Property Get Count() As Long: End Property

''
' Copies the contents of the collection to an array.
'
' @param arr The array to copy the elements into.
' @param index The index in the destination array to start copying the elements to.
'
Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long): End Sub
