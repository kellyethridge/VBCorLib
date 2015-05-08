VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodePageDecoder"
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
'    Module: CodePageDecoder
'

''
' This class is a general decoding wrapper around the Windows decoding APIs.
'
Option Explicit
Implements IObject
Implements Decoder

Private mCodePage       As Long
Private mMaxCharSize    As Long
Private mLastByte       As Long
Private mTemp(1)        As Byte



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal MaxCharSize As Long, ByVal CodePage As Long)
    mLastByte = -1
    mMaxCharSize = MaxCharSize
    mCodePage = CodePage
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BytesAreCharacterAligned(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Boolean
    Dim i As Long
    i = Index + Count - 1
    Do While i > Index
        If Not IsDBCSLeadByteEx(mCodePage, Bytes(i - 1)) Then Exit Do
        i = i - 1
    Loop
    BytesAreCharacterAligned = ((((Index + Count) - i) And 1) = 0)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Decoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = IObject_Equals(Value)
End Function

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    Dim Result As Long
    
    Result = VerifyArrayRange(SAPtr(Bytes), Index, Count)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", Index, "Index", Count, "Count"
    
    Dim Offset As Long
    If mLastByte >= 0 Then
        Index = Index + 1
        Count = Count - 1
        Offset = 1
        If Count = 0 Then
            Decoder_GetCharCount = 1
            Exit Function
        End If
    End If
    
    ' now check if we ended in the middle of a multi-byte
    ' character and if so, don't count it.
    If Not BytesAreCharacterAligned(Bytes, Index, Count) Then
        Count = Count - 1
        If Count = 0 Then
            Decoder_GetCharCount = Offset
            Exit Function
        End If
    End If
    
    Decoder_GetCharCount = Offset + MultiByteToWideChar(mCodePage, 0, Bytes(Index), Count, ByVal 0&, 0)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim Result As Long
    
    Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
    If Result <> NO_ERROR Then _
        ThrowArrayRangeException Result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount"
    If CorArray.IsNull(Chars) Then _
        Throw Cor.NewArgumentNullException("Chars", Environment.GetResourceString(ArgumentNull_Array))
    If CharIndex < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException("CharIndex", Message:=Environment.GetResourceString(ArgumentOutOfRange_LBound))
    
    If ByteCount = 0 Then Exit Function
    
    Dim Offset As Long
    If mLastByte >= 0 Then
        If CharIndex > UBound(Chars) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
        
        mTemp(0) = mLastByte
        mTemp(1) = Bytes(ByteIndex)
        Call MultiByteToWideChar(mCodePage, 0, mTemp(0), 2, Chars(CharIndex), 1)
        ByteIndex = ByteIndex + 1
        ByteCount = ByteCount - 1
        CharIndex = CharIndex + 1
        mLastByte = -1
        Offset = 1
        If ByteCount = 0 Then
            Decoder_GetChars = 1
            Exit Function
        End If
    End If
    
    If mMaxCharSize > 1 Then
        If Not BytesAreCharacterAligned(Bytes, ByteIndex, ByteCount) Then
            ByteCount = ByteCount - 1
            mLastByte = Bytes(ByteIndex + ByteCount)
            If ByteCount = 0 Then
                Decoder_GetChars = Offset
                Exit Function
            End If
        End If
    End If

    If CharIndex > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
    
    Decoder_GetChars = Offset + MultiByteToWideChar(mCodePage, 0, Bytes(ByteIndex), ByteCount, Chars(CharIndex), UBound(Chars) - CharIndex + 1)
End Function

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = IObject_GetHashcode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = IObject_ToString
End Function
