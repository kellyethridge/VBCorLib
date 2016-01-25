VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SafeHandle"
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
' Module: SafeHandle
'
Option Explicit

Private mHandle         As Long
Private mIsClosed       As Boolean
Private mOwnsHandle     As Boolean
Private mZeroIsInvalid  As Boolean

Public Event ReleaseHandle()


''
' Returns the wrapped handle.
'
Public Function DangerousGetHandle() As Long
Attribute DangerousGetHandle.VB_UserMemId = 0
    DangerousGetHandle = mHandle
End Function

''
' Indicates if the SafeHandle is considered closed.
'
Public Property Get IsClosed() As Boolean
    IsClosed = mIsClosed
End Property

''
' Indicates if the wrapped handle is considered invalid.
'
' @return Returns true if the handle is invalide, false otherwise.
' @remarks An invalid handle is equal to the value -1.
'
Public Property Get IsInvalid() As Boolean
    IsInvalid = (mHandle = INVALID_HANDLE) Or (mZeroIsInvalid And (mHandle = 0))
End Property

''
' Attempts to close the handle by raising the ReleaseHandle event.
'
' @remarks An event is raise indicating the handle should be closed by
' a subclass. The reference to this class must use WithEvents to correctly
' conduct the closing of the handle.
'
Public Sub CloseHandle()
    If mOwnsHandle And Not mIsClosed Then
        RaiseEvent ReleaseHandle
        mHandle = INVALID_HANDLE
        mIsClosed = True
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal PreexistingHandle As Long, ByVal OwnsHandle As Boolean, ByVal ZeroIsInvalid As Boolean)
    mHandle = PreexistingHandle
    mOwnsHandle = OwnsHandle
    mZeroIsInvalid = ZeroIsInvalid
End Sub


