VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SafeFindFileHandle"
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
' Module: FindFileHandle
'
Option Explicit
Implements ISafeHandle

Private mFileHandle As Long
Private mIsClosed   As Boolean

Public Property Get IsValid() As Boolean
    IsValid = (mFileHandle <> INVALID_HANDLE)
End Property

Public Property Get DangerousHandle() As Long
    DangerousHandle = mFileHandle
End Property

Public Sub CloseHandle()
    FindClose mFileHandle
    mIsClosed = True
End Sub

Public Property Get IsClosed() As Boolean
    IsClosed = mIsClosed
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Initialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal FileHandle As Long)
    mFileHandle = FileHandle
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    CloseHandle
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ISafeHandle Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ISafeHandle_Close()
    CloseHandle
End Sub

Private Property Get ISafeHandle_DangerousHandle() As Long
    ISafeHandle_DangerousHandle = DangerousHandle
End Property

Private Property Get ISafeHandle_IsClosed() As Boolean
    ISafeHandle_IsClosed = IsClosed
End Property

Private Property Get ISafeHandle_IsValid() As Boolean
    ISafeHandle_IsValid = IsValid
End Property

Private Property Get ISafeHandle_This() As VBCorType.ISafeHandle
    Set ISafeHandle_This = Me
End Property
