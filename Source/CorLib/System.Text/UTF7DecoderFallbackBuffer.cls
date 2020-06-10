VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7DecoderFallbackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: UTF7DecoderFallbackBuffer
'
Option Explicit
Implements IObject
Implements DecoderFallbackBuffer

Private mFallbackChar   As Integer
Private mCount          As Long
Private mSize           As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mCount = -1
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DecoderFallbackBuffer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DecoderFallbackBuffer_Equals(Value As Variant) As Boolean
    DecoderFallbackBuffer_Equals = IObject_Equals(Value)
End Function

Private Function DecoderFallbackBuffer_Fallback(BytesUnknown() As Byte, ByVal Index As Long) As Boolean
    mFallbackChar = BytesUnknown(0)
    
    If mFallbackChar > 0 Then
        mCount = 1
        mSize = 1
        DecoderFallbackBuffer_Fallback = True
    End If
End Function

Private Function DecoderFallbackBuffer_GetHashCode() As Long
    DecoderFallbackBuffer_GetHashCode = IObject_GetHashCode
End Function

Private Function DecoderFallbackBuffer_GetNextChar() As Integer
    If mCount > 0 Then
        mCount = mCount - 1
        DecoderFallbackBuffer_GetNextChar = mFallbackChar
    End If
End Function

Private Function DecoderFallbackBuffer_MovePrevious() As Boolean
    If mCount >= 0 Then
        mCount = mCount + 1
    End If
    
    DecoderFallbackBuffer_MovePrevious = (mCount >= 0 And mCount <= mSize)
End Function

Private Property Get DecoderFallbackBuffer_Remaining() As Long
    If mCount > 0 Then
        DecoderFallbackBuffer_Remaining = mCount
    End If
End Property

Private Sub DecoderFallbackBuffer_Reset()
    mCount = -1
End Sub

Private Function DecoderFallbackBuffer_ToString() As String
    DecoderFallbackBuffer_ToString = IObject_ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function
