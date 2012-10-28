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
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: UTF7Decoder
'
Option Explicit
Implements IObject
Implements Decoder

Private mEncoding As UTF7Encoding

' state for GetCharCount

' state for GetChars
Private mInBase64   As Boolean
Private mBits       As Long
Private mBitCount   As Long
Private mHasPlus    As Boolean


Friend Sub Init(ByVal Encoding As UTF7Encoding)
    Set mEncoding = Encoding
End Sub

Friend Property Get InBase64() As Boolean
    InBase64 = mInBase64
End Property

Friend Property Let InBase64(ByVal RHS As Boolean)
    mInBase64 = RHS
End Property

Friend Property Get bits() As Currency
    bits = mBits
End Property

Friend Property Let bits(ByVal RHS As Currency)
    mBits = RHS
End Property

Friend Property Get BitCount() As Long
    BitCount = mBitCount
End Property

Friend Property Let BitCount(ByVal RHS As Long)
    mBitCount = RHS
End Property

Friend Property Get HasPlus() As Boolean
    HasPlus = mHasPlus
End Property

Friend Property Let HasPlus(ByVal RHS As Boolean)
    mHasPlus = RHS
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
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Bytes), Index, Count)
    If result <> NO_ERROR Then ThrowArrayRangeException result, "Bytes", Index, "Index", Count, "Count"
    
    Decoder_GetCharCount = mEncoding.InternalGetCharCount(Bytes, Index, Count)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
    If result <> NO_ERROR Then ThrowArrayRangeException result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount"
    
    Decoder_GetChars = mEncoding.InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Me)
End Function

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = IObject_GetHashcode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = IObject_ToString
End Function
