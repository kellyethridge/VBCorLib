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

Friend Property Let BytesLeft(ByVal RHS As Long)
    mBytesLeft = RHS
End Property

Friend Property Get Char() As Integer
    Char = mChar
End Property

Friend Property Let Char(ByVal RHS As Integer)
    mChar = RHS
End Property

Friend Property Get BytesInSequence() As Long
    BytesInSequence = mBytesInSequence
End Property

Friend Property Let BytesInSequence(ByVal RHS As Long)
    mBytesInSequence = RHS
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_ToString() As String
    IObject_ToString = Object.ToString(Me, App)
End Function

Private Function IObject_Equals(ByRef Value As Variant) As Boolean
    IObject_Equals = Object.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Decoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = IObject_Equals(Value)
End Function

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    Decoder_GetCharCount = mEncoding.GetCharCount(Bytes, Index, Count)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
    If result <> NO_ERROR Then Call ThrowArrayRangeException(result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount")
    
    Decoder_GetChars = mEncoding.InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Me)
End Function

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = IObject_GetHashcode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = IObject_ToString
End Function
