VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumeratorBase"
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
' Module: EnumeratorBase
'

''
' Used to manage an enumeration index into a set of items.
'
' @remarks Use this class to handle the starting, advancing and finishing of an index.
' <p>The class itself does not access a list of items. Instead, it manages an index into a list of items
' that are being enumerated. It deals with ensuring the enumeration starts correctly and when
' the enumeration has finished.</p>
' <p>To create a new instance use the <a href="Constructors.NewEnumeratorBase.html">NewEnumeratorBase</a> method.</p>
' @see Constructors
Option Explicit
Implements ICloneable

Private mLength         As Long
Private mCount          As Long
Private mStartIndex     As Long
Private mCollection     As IVersionable
Private mVersion        As Long


''
' Advances the current index, returning if the advancement succeeded.
'
' @return <b>True</b> if the index was advanced; otherwise <b>False</b>.
' @remarks <b>MoveNext</b> will return <b>False</b> when the enumeration
' initially finishes. If additional calls are made to <b>MoveNext</b> an
' <a href="InvalidOperationException.html">InvalidOperationException</a>
' will be thrown.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The enumeration has already finished.</td>
'   </tr>
' </table>
Public Function MoveNext() As Boolean
    EnsureUnchangedVersion
    EnsureNotFinished
    
    mCount = mCount + 1
    MoveNext = mCount < mLength
End Function

''
' Resets the index back to the initial starting position.
Public Sub Reset()
    mCount = -1
End Sub

''
' Gets the current index.
'
' @return The current index.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td>
'        <td>
'           Attempted to get the current index before enumeration has started.<br/>
'           -or-<br/>
'           Attempted to get the current index after enumeration has finished.<br/>
'           -or-<br/>
'           The underlying collection has changed.
'        </td>
'   </tr>
' </table>
Public Property Get Index() As Long
    EnsureUnchangedVersion
    EnsureStarted
    EnsureNotFinished
    
    Index = mStartIndex + mCount
End Property

''
' Creates a new instance of the enumerator base at the current index position.
'
' @return Returns a clone of the current enumerator base class.
Public Function Clone() As EnumeratorBase
    Set Clone = New EnumeratorBase
    Clone.InitClone mStartIndex, mLength, mCount, mCollection
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal StartIndex As Long, ByVal Length As Long, ByVal Collection As IVersionable)
    If Length < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NeedNonNegNum
    
    mLength = Length
    mStartIndex = StartIndex
    mCount = -1
    Set mCollection = IfObject(Collection, NullVersionable.Instance)
    mVersion = mCollection.Version
End Sub

Friend Sub InitClone(ByVal StartIndex As Long, ByVal Length As Long, ByVal Count As Long, ByVal Collection As IVersionable)
    mStartIndex = StartIndex
    mLength = Length
    mCount = Count
    Set mCollection = Collection
    mVersion = mCollection.Version
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureUnchangedVersion()
    If mCollection.Version <> mVersion Then
        Error.InvalidOperation InvalidOperation_VersionError
    End If
End Sub

Private Sub EnsureStarted()
    If mCount < 0 Then
        Error.InvalidOperation InvalidOperation_EnumNotStarted
    End If
End Sub

Private Sub EnsureNotFinished()
    If mCount >= mLength Then
        Error.InvalidOperation InvalidOperation_EnumFinished
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function
