VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultDecoder"
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
' Module: DefaultDecoder
'

Option Explicit
Implements IObject
Implements Decoder

Private mEncoding       As Encoding
Private mFallback       As DecoderFallback
Private mFallbackBuffer As DecoderFallbackBuffer



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Encoding As Encoding)
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
    If Not mFallbackBuffer Is Nothing Then
        If mFallbackBuffer.Remaining > 0 Then
            Error.Argument Argument_FallbackBufferNotEmpty, "Value"
        End If
    End If
    
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
    Decoder_GetCharCount = mEncoding.GetCharCount(Bytes, Index, Count)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long, Optional ByVal Flush As Boolean) As Long
    ValidateDecoderGetChars Bytes, ByteIndex, ByteCount, Chars, CharIndex
    Decoder_GetChars = mEncoding.GetCharsEx(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
End Function

Private Sub Decoder_Reset()
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
