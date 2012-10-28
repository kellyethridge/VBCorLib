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
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: StreamAsyncResult
'
''
' Class derived from IObject.
'

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

Friend Property Set Exception(ByVal RHS As Exception)
    Set mException = RHS
End Property

Friend Property Get IsReadType() As Boolean
    IsReadType = mIsReadType
End Property

Friend Property Let IsReadType(ByVal RHS As Boolean)
    mIsReadType = RHS
End Property

Friend Property Get EndCalled() As Boolean
    EndCalled = mEndCalled
End Property

Friend Property Let EndCalled(ByVal RHS As Boolean)
    mEndCalled = RHS
End Property

Friend Property Get CompletedSynchronously() As Boolean
    CompletedSynchronously = mCompletedSynchronously
End Property

Friend Property Let CompletedSynchronously(ByVal RHS As Boolean)
    mCompletedSynchronously = RHS
End Property

Friend Property Get BytesRead() As Long
    BytesRead = mBytesRead
End Property

Friend Property Let BytesRead(ByVal RHS As Long)
    mBytesRead = RHS
End Property

Friend Property Get IsCompleted() As Boolean
    IsCompleted = mIsCompleted
End Property

Friend Property Let IsCompleted(ByVal RHS As Boolean)
    mIsCompleted = RHS
End Property


''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Friend Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Friend Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Friend Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


Friend Sub Init(ByVal ObjectState As Variant)
    Helper.MoveVariant mObjectState, ObjectState
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
