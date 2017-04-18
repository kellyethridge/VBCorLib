VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8Encoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: UTF8Encoder
'

''
' Private class returned by UTF8Encoding.GetEncoder
'

Option Explicit
Implements IObject
Implements Encoder

Private mEncoding           As UTF8Encoding
Private mHighSurrogateChar  As Long
Private mFlush              As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Encoding As UTF8Encoding)
    Set mEncoding = Encoding
End Sub

Friend Property Get Flush() As Boolean
    Flush = mFlush
End Property

Friend Property Get HighSurrogateChar() As Long
    HighSurrogateChar = mHighSurrogateChar
End Property

Friend Property Let HighSurrogateChar(ByVal RHS As Long)
    mHighSurrogateChar = RHS
End Property


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

End Sub

Private Property Set Encoder_Fallback(ByVal RHS As EncoderFallback)

End Property

Private Property Get Encoder_Fallback() As EncoderFallback

End Property

Private Property Get Encoder_FallbackBuffer() As EncoderFallbackBuffer

End Property

Private Function Encoder_Equals(Value As Variant) As Boolean
    Encoder_Equals = IObject_Equals(Value)
End Function

Private Function Encoder_GetByteCount(Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Flush As Boolean) As Long
    mFlush = Flush
    
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Chars), Index, Count)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Chars", Index, "Index", Count, "Count"
    
    Encoder_GetByteCount = mEncoding.GetBytesCore(Chars, Index, Count, NullBytes, 0, Me)
End Function

Private Function Encoder_GetBytes(Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long, ByVal Flush As Boolean) As Long
    mFlush = Flush
    
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Chars), CharIndex, CharCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Chars", CharIndex, "CharIndex", CharCount, "CharCount")
    
    Encoder_GetBytes = mEncoding.InternalGetBytes(Chars, CharIndex, CharCount, Bytes, ByteIndex, Me)
End Function

Private Function Encoder_GetHashCode() As Long
    Encoder_GetHashCode = IObject_GetHashCode
End Function

Private Sub Encoder_Reset()

End Sub

Private Function Encoder_ToString() As String
    Encoder_ToString = IObject_ToString
End Function
