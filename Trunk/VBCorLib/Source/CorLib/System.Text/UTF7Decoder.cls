VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7Decoder"
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
' Module: UTF7Decoder
'
Option Explicit
Implements IObject
Implements Decoder

Private mEncoding   As UTF7Encoding
Private mInBase64   As Boolean
Private mBits       As Long
Private mBitCount   As Long
Private mHasPlus    As Boolean


Friend Property Get InBase64() As Boolean
    InBase64 = mInBase64
End Property

Friend Property Let InBase64(ByVal Value As Boolean)
    mInBase64 = Value
End Property

Friend Property Get Bits() As Currency
    Bits = mBits
End Property

Friend Property Let Bits(ByVal Value As Currency)
    mBits = Value
End Property

Friend Property Get BitCount() As Long
    BitCount = mBitCount
End Property

Friend Property Let BitCount(ByVal Value As Long)
    mBitCount = Value
End Property

Friend Property Get HasPlus() As Boolean
    HasPlus = mHasPlus
End Property

Friend Property Let HasPlus(ByVal Value As Boolean)
    mHasPlus = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
'   Decoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = IObject_Equals(Value)
End Function
Private Property Set Decoder_Fallback(ByVal RHS As DecoderFallback)

End Property

Private Property Get Decoder_Fallback() As DecoderFallback

End Property

Private Property Get Decoder_FallbackBuffer() As DecoderFallbackBuffer

End Property

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, Optional ByVal Flush As Boolean) As Long
    Decoder_GetCharCount = mEncoding.GetCharCount(Bytes, Index, Count)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long, Optional ByVal Flush As Boolean) As Long
    Dim CharCount As Long
    
    Encoding.ValidateByteRange Bytes, ByteIndex, ByteCount
    CharCount = mEncoding.InternalGetCharCount(Bytes, ByteIndex, ByteCount)
    Encoding.ValidateCharRange Chars, CharIndex, CharCount
    
    Decoder_GetChars = mEncoding.InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Me)
End Function

Private Sub Decoder_Reset()

End Sub

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = IObject_GetHashCode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = IObject_ToString
End Function
