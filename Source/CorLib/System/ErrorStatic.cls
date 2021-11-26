VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorStatic"
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
' Module: ErrorStatic
'
Option Explicit

Friend Sub Argument(ByVal Key As ResourceStringKey, Optional ByRef ParamName As String)
    Throw Cor.NewArgumentException(Environment.GetResourceString(Key), ParamName)
End Sub

Friend Sub ArgumentNull(ByRef ParamName As String, Optional ByVal Key As ResourceStringKey = ArgumentNull_Generic)
    Throw Cor.NewArgumentNullException(ParamName, Environment.GetResourceString(Key))
End Sub

Friend Sub ArgumentOutOfRange(ByRef ParamName As String, Optional ByVal Key As ResourceStringKey = ArgumentOutOfRange_Exception)
    Throw Cor.NewArgumentOutOfRangeException(ParamName, Environment.GetResourceString(Key))
End Sub

Friend Sub Rank(Optional ByVal Key As ResourceStringKey = Rank_MultiDimNotSupported)
    Throw Cor.NewRankException(Environment.GetResourceString(Key))
End Sub

Friend Sub ObjectDisposed(Optional ByVal Key As ResourceStringKey = ObjectDisposed_Generic)
    Throw Cor.NewObjectDisposedException(vbNullString, Environment.GetResourceString(Key))
End Sub

Friend Sub IndexOutOfRange(ByVal Key As ResourceStringKey)
    Throw Cor.NewIndexOutOfRangeException(Environment.GetResourceString(Key))
End Sub

Friend Sub InvalidOperation(ByVal Key As ResourceStringKey)
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(Key))
End Sub

Friend Sub NotSupported(ByVal Key As ResourceStringKey)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(Key))
End Sub

Friend Sub IO(ByVal Key As ResourceStringKey)
    Throw Cor.NewIOException(Environment.GetResourceString(Key))
End Sub

Friend Sub EndOfStream(ByVal Key As ResourceStringKey)
    Throw Cor.NewEndOfStreamException(Environment.GetResourceString(Key))
End Sub

Friend Sub Win32Error(ByVal ErrorCode As Long)
    Dim Message As String
    Message = GetMessage(ErrorCode)
    
    Select Case ErrorCode
        Case ERROR_ACCESS_DENIED
            Throw Cor.NewUnauthorizedAccessException(Message)
        Case Else
            Throw Cor.NewApplicationException(Message)
    End Select
End Sub

Friend Sub Format(ByVal Key As ResourceStringKey)
    Throw Cor.NewFormatException(Environment.GetResourceString(Key))
End Sub

Friend Sub Cryptographic(ByVal Key As ResourceStringKey)
    Dim Message As String
    
    If Key = LastDllError Then
        Message = GetMessage(Err.LastDllError)
    Else
        Message = Environment.GetResourceString(Key)
    End If
    
    Throw Cor.NewCryptographicException(Message)
End Sub

Friend Sub UnauthorizedAccess(ByVal Key As ResourceStringKey)
    Throw Cor.NewUnauthorizedAccessException(Environment.GetResourceString(Key))
End Sub

Friend Sub Overflow(ByVal Key As ResourceStringKey)
    Throw Cor.NewOverflowException(Environment.GetResourceString(Key))
End Sub
