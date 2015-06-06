VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: ArrayListStatic
'

''
' Represents the static methods of the ArrayList class.
'
' @remarks This class cannot be directly instantiated. In order to access
' the method of this class, use the <b>ArrayList.*</b> syntax.
' <pre>
' Set ReadOnlyList = ArrayList.ReadOnly(list)
' </pre>
' @see ArrayList
'
Option Explicit

''
' Returns an ArrayList object wrapped inside a readonly object.
'
' @param list The ArrayList to wrap
' @return The readonly ArrayList wrapper.
' @remarks Passing the object around provides access to the underlying
' ArrayList, however, the list cannot be changed in size or content.
'
Public Function ReadOnly(ByVal List As ArrayList) As ArrayList
    If List Is Nothing Then _
        Error.ArgumentNull "List"
    
    Dim Ret As New ReadOnlyArrayList
    Ret.Init List
    Set ReadOnly = Ret
End Function

''
' Returns an ArrayList object that is wrapped in a fixedsize object.
'
' @param list The ArrayList to be wrapped.
' @return The wrapped ArrayList.
' @remarks Passing this object around provides access to the underlying
' ArrayList, however, prevents the list from changing size. The current
' values in the list can be modified, but the size of the list cannot change.
'
Public Function FixedSize(ByVal List As ArrayList) As ArrayList
    If List Is Nothing Then _
        Throw Cor.NewArgumentNullException("List")
    
    Dim Ret As New FixedSizeArrayList
    Ret.Init List
    Set FixedSize = Ret
End Function

Public Function FixedSizeList(ByVal List As IList) As IList

End Function

''
' Creates a new ArrayList and fills it with the supplied value the specified number of times.
'
' @param Value The value to fill the ArrayList with.
' @param Count The number of times the value is to be added to the list.
' @return A new ArrayList.
'
Public Function Repeat(ByRef Value As Variant, ByVal Count As Long) As ArrayList
    If Count < 0 Then _
        Throw Error.ArgumentOutOfRange("Count", ArgumentOutOfRange_NeedNonNegNum)
    
    Set Repeat = New ArrayList
    Repeat.Capacity = Count
    
    Dim i As Long
    For i = 1 To Count
        Repeat.Add Value
    Next i
End Function

''
' Creates an object that allows any IList object to be treated as an ArrayList instance.
'
' @param list The list to be wrapped as an ArrayList instance.
' @return A new object that can be used like an ArrayList
' @remarks The adapter gives all IList objects the ability to be used anywhere an
' ArrayList instance would be used instead.
'
Public Function Adapter(ByVal List As IList) As ArrayList
    If List Is Nothing Then _
        Error.ArgumentNull "List"
    
    Dim Wrapper As New IListWrapper
    Wrapper.Init List
    Set Adapter = Wrapper
End Function
