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
' Represents the static methods of the <a href="ArrayList.html">ArrayList</a> class.
'
' @remarks This class cannot be directly instantiated. In order to access
' the method of this class, use the <b>ArrayList.<i>Method</i></b> syntax.
' <pre>
' Set ReadOnlyList = ArrayList.ReadOnly(list)
' </pre>
' @see ArrayList
'
Option Explicit

''
' Returns an <a href="ArrayList.html">ArrayList</a> object wrapped inside a read-only list.
'
' @param List The <a href="ArrayList.html">ArrayList</a> to wrap.
' @return The read-only <a href="ArrayList.html">ArrayList</a> wrapper around <i>List</i>.
' @remarks Passing the object around provides access to the underlying
' <a href="ArrayList.html">ArrayList</a>, however, the list cannot be changed in size or content.
Public Function ReadOnly(ByVal List As ArrayList) As ArrayList
    If List Is Nothing Then _
        Error.ArgumentNull "List"
    
    Dim Result As New ReadOnlyArrayList
    Result.Init List
    Set ReadOnly = Result
End Function

''
' Returns an <a href="ArrayList.html">ArrayList</a> object that is wrapped in a fixed-size object.
'
' @param list The <a href="ArrayList.html">ArrayList</a> to be wrapped.
' @return The wrapped <a href="ArrayList.html">ArrayList</a>.
' @remarks Passing this object around provides access to the underlying
' <a href="ArrayList.html">ArrayList</a>, however, prevents the list from changing size. The current
' values in the list can be modified, but the size of the list cannot change.
Public Function FixedSize(ByVal List As ArrayList) As ArrayList
    If List Is Nothing Then _
        Error.ArgumentNull "List"
    
    Dim Ret As New FixedSizeArrayList
    Ret.Init List
    Set FixedSize = Ret
End Function

''
' Creates a new <a href="ArrayList.html">ArrayList</a> and fills it with the supplied value the specified number of times.
'
' @param Value The value to fill the <a href="ArrayList.html">ArrayList</a> with.
' @param Count The number of times the value is to be added to the list.
' @return An <a href="ArrayList.html">ArrayList</a> with <i>Count</i> elements, all of which are copies of <i>Value</i>.
Public Function Repeat(ByRef Value As Variant, ByVal Count As Long) As ArrayList
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    Set Repeat = New ArrayList
    Repeat.Capacity = Count
    
    Dim i As Long
    For i = 1 To Count
        Repeat.Add Value
    Next i
End Function

''
' Creates an object that allows any <a href="IList.html">IList</a> object to be treated as an <a href="ArrayList.html">ArrayList</a> instance.
'
' @param List The list to be wrapped as an <a href="ArrayList.html">ArrayList</a> instance.
' @return An <a href="IList.html">IList</a> object that can be used like an <a href="ArrayList.html">ArrayList</a>.
' @remarks The adapter gives all <a href="IList.html">IList</a> objects the ability to be used anywhere an
' <a href="ArrayList.html">ArrayList</a> instance would be used instead.
Public Function Adapter(ByVal List As IList) As ArrayList
    If List Is Nothing Then _
        Error.ArgumentNull "List"
    
    Dim Wrapper As New IListWrapper
    Wrapper.Init List
    Set Adapter = Wrapper
End Function
