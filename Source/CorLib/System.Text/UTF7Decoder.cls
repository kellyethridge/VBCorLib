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

Private mEncoding       As UTF7Encoding
Private mFallback       As DecoderFallback
Private mFallbackBuffer As DecoderFallbackBuffer
Private mInBase64       As Boolean
Private mBits           As Long
Private mBitCount       As Long
Private mHasPlus        As Boolean
Private mFlush          As Boolean


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

Friend Property Get Flush() As Boolean
    Flush = mFlush
End Property

Friend Property Let Flush(ByVal Value As Boolean)
    mFlush = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Encoding As UTF7Encoding)
    Set mEncoding = Encoding
    Set mFallback = Encoding.DecoderFallback
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
'   Decoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Decoder_Convert(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByVal Flush As Boolean, BytesUsed As Long, CharsUsed As Long, Completed As Boolean)
    DecoderConvert Me, Bytes, ByteIndex, ByteCount, Chars, CharIndex, CharCount, Flush, BytesUsed, CharsUsed, Completed
End Sub

Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = IObject_Equals(Value)
End Function
Private Property Set Decoder_Fallback(ByVal Value As DecoderFallback)
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Set mFallback = Value
    Set mFallbackBuffer = Nothing
End Property

Private Property Get Decoder_Fallback() As DecoderFallback
    Set Decoder_Fallback = mFallback
End Property

Private Property Get Decoder_FallbackBuffer() As DecoderFallbackBuffer
    If mFallbackBuffer Is Nothing Then
        Set mFallbackBuffer = mFallback.CreateFallbackBuffer
    End If
    
    Set Decoder_FallbackBuffer = mFallbackBuffer
End Property

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, Optional ByVal Flush As Boolean) As Long
    ValidateDecoderGetCharCount Bytes, Index, Count
    
    mFlush = Flush
    Decoder_GetCharCount = mEncoding.GetCharCountCore(Bytes, Index, Count, Me)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long, Optional ByVal Flush As Boolean) As Long
    ValidateDecoderGetChars Bytes, ByteIndex, ByteCount, Chars, CharIndex
    
    mFlush = Flush
    Decoder_GetChars = mEncoding.GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Me)
End Function

Private Sub Decoder_Reset()
    mInBase64 = False
    mBits = 0
    mBitCount = 0
    mHasPlus = False
    mFlush = False
    
    If Not mFallbackBuffer Is Nothing Then
        mFallbackBuffer.Reset
    End If
End Sub

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = IObject_GetHashCode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = IObject_ToString
End Function
