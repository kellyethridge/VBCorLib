VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SafeFileHandle"
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
' Module: SafeFileHandle
'
Option Explicit
Implements SafeHandle

Private WithEvents mBase As SafeHandle
Attribute mBase.VB_VarHelpID = -1


Public Property Get IsInvalid() As Boolean
    IsInvalid = mBase.IsInvalid
End Property

Public Function DangerousGetHandle() As Long
Attribute DangerousGetHandle.VB_UserMemId = 0
    DangerousGetHandle = mBase.DangerousGetHandle
End Function

Public Sub CloseHandle()
    mBase.CloseHandle
End Sub

Public Property Get IsClosed() As Boolean
    IsClosed = mBase.IsClosed
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal PreexistingHandle As Long, ByVal OwnsHandle As Boolean)
    Set mBase = Cor.NewSafeHandle(PreexistingHandle, OwnsHandle)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    CloseHandle
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SafeHandle
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mBase_ReleaseHandle()
    CorType.CloseHandle DangerousGetHandle
End Sub

Private Sub SafeHandle_CloseHandle()
    CloseHandle
End Sub

Private Function SafeHandle_DangerousGetHandle() As Long
    SafeHandle_DangerousGetHandle = DangerousGetHandle
End Function

Private Property Get SafeHandle_IsClosed() As Boolean
    SafeHandle_IsClosed = IsClosed
End Property

Private Property Get SafeHandle_IsInvalid() As Boolean
    SafeHandle_IsInvalid = IsInvalid
End Property

