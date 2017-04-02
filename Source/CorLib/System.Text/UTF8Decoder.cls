VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8Decoder"
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
'    Module: UTF8Decoder
'

''
' Private class returned by UTF8Encoding.GetDecoder
'
Option Explicit
Implements IObject
Implements Decoder

Private mEncoding           As UTF8Encoding
Private mBytesLeft          As Long
Private mChar               As Long
Private mBytesInSequence    As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Encoding As UTF8Encoding)
    Set mEncoding = Encoding
End Sub

Friend Property Get BytesLeft() As Long
    BytesLeft = mBytesLeft
End Property

Friend Property Let BytesLeft(ByVal Value As Long)
    mBytesLeft = Value
End Property

Friend Property Get Char() As Integer
    Char = mChar
End Property

Friend Property Let Char(ByVal Value As Integer)
    mChar = Value
End Property

Friend Property Get BytesInSequence() As Long
    BytesInSequence = mBytesInSequence
End Property

Friend Property Let BytesInSequence(ByVal Value As Long)
    mBytesInSequence = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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

End Sub

Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = IObject_Equals(Value)
End Function

Private Property Set Decoder_Fallback(ByVal Value As DecoderFallback)

End Property

Private Property Get Decoder_Fallback() As DecoderFallback

End Property

Private Property Get Decoder_FallbackBuffer() As DecoderFallbackBuffer

End Property

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, Optional ByVal Flush As Boolean) As Long
    Decoder_GetCharCount = mEncoding.GetCharCount(Bytes, Index, Count)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long, Optional ByVal Flush As Boolean) As Long
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount")
    
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
