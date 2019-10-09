VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SafeMemoryMappedFileHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: SafeMemoryMappedFileHandle
'

''
' Provides a safe handle that represents a memory-mapped file for sequential access.
'
Option Explicit
Implements IObject
Implements SafeHandle

Private WithEvents mBase As SafeHandle
Attribute mBase.VB_VarHelpID = -1


''
' When overridden in a derived class, gets a value indicating whether the handle value is invalid.
'
' @return <b>True</b> if the handle value is invalid; otherwise, <b>False</b>.
Public Property Get IsInvalid() As Boolean
    IsInvalid = mBase.IsInvalid
End Property

''
' Returns the value of the handle field.
'
' @return An <b>Long</b> representing the value of the handle field.
Public Function DangerousGetHandle() As Long
    DangerousGetHandle = mBase.DangerousGetHandle
End Function

''
' Marks the handle for releasing and freeing resources.
Public Sub CloseHandle()
    mBase.CloseHandle
End Sub

''
' Gets a value indicating whether the handle is closed.
'
' @return <b>True</b> if the handle is closed; otherwise, <b>False</b>.
Public Property Get IsClosed() As Boolean
    IsClosed = mBase.IsClosed
End Property

''
' Determines whether the specified value passed in is equal to the current object.
'
' @param Value The value to compare to the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string that represents the current object.
'
' @return A string the represents the current object.
' @see IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal PreexistingHandle As Long)
    Set mBase = Cor.NewSafeHandle(PreexistingHandle, True, True)
End Sub

Private Sub Class_Terminate()
    CloseHandle
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


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
