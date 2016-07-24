VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7Encoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: UTF7Encoder
'
Option Explicit
Implements IObject
Implements Encoder

Private mEncoding   As UTF7Encoding
Private mFlush      As Boolean
Private mBits       As Long
Private mBitCount   As Long


Friend Property Get Flush() As Boolean
    Flush = mFlush
End Property

Friend Property Get Bits() As Long
    Bits = mBits
End Property

Friend Property Let Bits(ByVal Value As Long)
    mBits = Value
End Property

Friend Property Get BitCount() As Long
    BitCount = mBitCount
End Property

Friend Property Let BitCount(ByVal Value As Long)
    mBitCount = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mBitCount = -1
End Sub

Friend Sub Init(ByVal Encoding As UTF7Encoding)
    Set mEncoding = Encoding
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function

Private Function IObject_Equals(ByRef Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Encoder_Equals(Value As Variant) As Boolean
    Encoder_Equals = IObject_Equals(Value)
End Function

Private Function Encoder_GetByteCount(Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Flush As Boolean) As Long
    Encoding.ValidateCharRange Chars, Index, Count
    
    mFlush = Flush
    
    Encoder_GetByteCount = mEncoding.InternalGetByteCount(Chars, Index, Count, Me)
End Function

Private Function Encoder_GetBytes(Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long, ByVal Flush As Boolean) As Long
    Dim ByteCount As Long
    
    Encoding.ValidateCharRange Chars, CharIndex, CharCount
    ByteCount = mEncoding.InternalGetByteCount(Chars, CharIndex, CharCount, Me)
    Encoding.ValidateByteRange Bytes, ByteIndex, ByteCount
    
    mFlush = Flush
    
    Encoder_GetBytes = mEncoding.InternalGetBytes(Chars, CharIndex, CharCount, Bytes, ByteIndex, Me)
End Function

Private Function Encoder_GetHashCode() As Long
    Encoder_GetHashCode = IObject_GetHashCode
End Function

Private Function Encoder_ToString() As String
    Encoder_ToString = IObject_ToString
End Function
