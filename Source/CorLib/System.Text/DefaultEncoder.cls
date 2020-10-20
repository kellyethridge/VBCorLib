VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultEncoder"
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
' Module: DefaultEncoder
'
'@Folder("CorLib.System.Text")
Option Explicit
Implements IObject
Implements Encoder

Private mEncoding       As Encoding
Private mFallback       As EncoderFallback
Private mFallbackBuffer As EncoderFallbackBuffer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Encoding As Encoding)
    Set mEncoding = Encoding
    Set mFallback = Encoding.EncoderFallback
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
Private Sub Encoder_Convert(Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByVal Flush As Boolean, CharsUsed As Long, BytesUsed As Long, Completed As Boolean)
    EncoderConvert mEncoding, Chars, CharIndex, CharCount, Bytes, ByteIndex, ByteCount, Flush, CharsUsed, BytesUsed, Completed
End Sub

Private Property Set Encoder_Fallback(ByVal Value As EncoderFallback)
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

Private Property Get Encoder_Fallback() As EncoderFallback
    Set Encoder_Fallback = mFallback
End Property

Private Property Get Encoder_FallbackBuffer() As EncoderFallbackBuffer
    If mFallbackBuffer Is Nothing Then
        Set mFallbackBuffer = mFallback.CreateFallbackBuffer
    End If
    
    Set Encoder_FallbackBuffer = mFallbackBuffer
End Property

Private Function Encoder_Equals(Value As Variant) As Boolean
    Encoder_Equals = IObject_Equals(Value)
End Function

Private Function Encoder_GetByteCount(Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Flush As Boolean) As Long
    ValidateEncoderGetByteCount Chars, Index, Count
    Encoder_GetByteCount = mEncoding.GetByteCount(Chars, Index, Count)
End Function

Private Function Encoder_GetBytes(Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long, ByVal Flush As Boolean) As Long
    ValidateEncoderGetBytes Chars, CharIndex, CharCount, Bytes, ByteIndex
    Encoder_GetBytes = mEncoding.GetBytesEx(Chars, CharIndex, CharCount, Bytes, ByteIndex)
End Function

Private Function Encoder_GetHashCode() As Long
    Encoder_GetHashCode = IObject_GetHashCode
End Function

Private Sub Encoder_Reset()
    If Not mFallbackBuffer Is Nothing Then
        mFallbackBuffer.Reset
    End If
End Sub

Private Function Encoder_ToString() As String
    Encoder_ToString = IObject_ToString
End Function
