VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathHelper"
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
' Module: PathHelper
'
Option Explicit

Private mFinalBuffer()  As Integer
Private mPosition       As Long

Friend Sub Append(ByVal Ch As Integer)
    If mPosition > UBound(mFinalBuffer) Then
        ReDim Preserve mFinalBuffer(0 To mPosition * 2)
    End If

    mFinalBuffer(mPosition) = Ch
    mPosition = mPosition + 1
End Sub

Friend Property Get Chars(ByVal Index As Long) As Integer
    Debug.Assert Index >= 0 And Index < mPosition
    
    Chars = mFinalBuffer(Index)
End Property

Friend Function GetFullPathName() As Long
    Dim NewBuffer() As Integer
    Dim NewLength   As Long
    
    ReDim NewBuffer(0 To mPosition + MAX_PATH)
    NewLength = GetFullPathNameW(mFinalBuffer(0), mPosition + MAX_PATH, NewBuffer(0), ByVal vbNullPtr)
    Helper.Swap4 ByVal ArrPtr(mFinalBuffer), ByVal ArrPtr(NewBuffer)
    mPosition = NewLength
    GetFullPathName = NewLength
End Function

Friend Property Get Length() As Long
    Length = mPosition
End Property

Friend Function ToString() As String
    ToString = SysAllocStringLen(VarPtr(mFinalBuffer(0)), mPosition)
End Function

Private Sub Class_Initialize()
    ReDim mFinalBuffer(0 To MAX_PATH)
End Sub
