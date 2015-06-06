VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionEnumeratorBase"
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
' Module: CollectionEnumeratorBase
'
Option Explicit
Implements ICloneable

Private mCollection As IVersionableCollection
Private mVersion    As Long
Private mLength     As Long
Private mCount      As Long
Private mStartIndex As Long


Public Sub CheckCurrent()
    EnsureUnchangedVersion
    EnsureStarted
    EnsureNotFinished
End Sub

Public Function MoveNext() As Boolean
    EnsureUnchangedVersion
    EnsureNotFinished
    
    mCount = mCount + 1
    MoveNext = mCount < mLength
End Function

Public Sub Reset()
    mCount = -1
End Sub

Public Property Get Index() As Long
    EnsureUnchangedVersion
    EnsureStarted
    EnsureNotFinished
    
    Index = mStartIndex + mCount
End Property

Public Function Clone() As CollectionEnumeratorBase
    Set Clone = New CollectionEnumeratorBase
    Clone.InitClone mCollection, mStartIndex, mLength, mCount
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Collection As IVersionableCollection, ByVal StartIndex As Long, ByVal Length As Long)
    If Collection Is Nothing Then _
        Throw Error.ArgumentNull("Collection")
    If Length < 0 Then _
        Throw Error.ArgumentOutOfRange("Length", ArgumentOutOfRange_NeedNonNegNum)
    
    Set mCollection = Collection
    mVersion = Collection.CurrentVersion
    mLength = Length
    mStartIndex = StartIndex
    Reset
End Sub

Friend Sub InitClone(ByVal Collection As IVersionableCollection, ByVal StartIndex As Long, ByVal Length As Long, ByVal Count As Long)
    Set mCollection = Collection
    mVersion = Collection.CurrentVersion
    mStartIndex = StartIndex
    mLength = Length
    mCount = Count
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureUnchangedVersion()
    If mCollection.CurrentVersion <> mVersion Then
        Throw Error.InvalidOperation(InvalidOperation_VersionError)
    End If
End Sub

Private Sub EnsureStarted()
    If mCount < 0 Then
        Throw Error.InvalidOperation(InvalidOperation_EnumNotStarted)
    End If
End Sub

Private Sub EnsureNotFinished()
    If mCount >= mLength Then
        Throw Error.InvalidOperation(InvalidOperation_EnumFinished)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

