VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamAsyncResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: StreamAsyncResult
'
'@Folder("CorLib.System.IO")
Option Explicit
Implements IObject
Implements IAsyncResult


Private mObjectState            As Variant
Private mBytesRead              As Long
Private mIsCompleted            As Boolean
Private mCompletedSynchronously As Boolean
Private mEndCalled              As Boolean
Private mIsReadType             As Boolean
Private mException              As Exception


Friend Property Get Exception() As Exception
    Set Exception = mException
End Property

Friend Property Set Exception(ByVal Value As Exception)
    Set mException = Value
End Property

Friend Property Get IsReadType() As Boolean
    IsReadType = mIsReadType
End Property

Friend Property Let IsReadType(ByVal Value As Boolean)
    mIsReadType = Value
End Property

Friend Property Get EndCalled() As Boolean
    EndCalled = mEndCalled
End Property

Friend Property Let EndCalled(ByVal Value As Boolean)
    mEndCalled = Value
End Property

Friend Property Get CompletedSynchronously() As Boolean
    CompletedSynchronously = mCompletedSynchronously
End Property

Friend Property Let CompletedSynchronously(ByVal Value As Boolean)
    mCompletedSynchronously = Value
End Property

Friend Property Get BytesRead() As Long
    BytesRead = mBytesRead
End Property

Friend Property Let BytesRead(ByVal Value As Long)
    mBytesRead = Value
End Property

Friend Property Get IsCompleted() As Boolean
    IsCompleted = mIsCompleted
End Property

Friend Property Let IsCompleted(ByVal Value As Boolean)
    mIsCompleted = Value
End Property


''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Friend Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Friend Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Friend Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef ObjectState As Variant)
    VariantCopyInd mObjectState, ObjectState
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IAsyncResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IAsyncResult_AsyncState() As Variant
    VariantCopy IAsyncResult_AsyncState, mObjectState
End Property

Private Property Get IAsyncResult_AsyncWaitHandle() As WaitHandle

End Property

Private Property Get IAsyncResult_CompletedSyncronously() As Boolean
    IAsyncResult_CompletedSyncronously = CompletedSynchronously
End Property

Private Property Get IAsyncResult_IsCompleted() As Boolean
    IAsyncResult_IsCompleted = IsCompleted
End Property
