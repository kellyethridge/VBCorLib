VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeDecoder"
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
'    Module: UnicodeDecoder
'
Option Explicit
Implements IObject
Implements Decoder

Private Const CHAR_SIZE As Long = 2


Private mIsBigEndian    As Boolean
Private mLastByte       As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal IsBigEndian As Boolean)
    mIsBigEndian = IsBigEndian
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InternalGetCharCount(ByVal Count As Long) As Long
    If mLastByte >= 0 Then Count = Count + 1
    InternalGetCharCount = Count \ CHAR_SIZE
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mLastByte = -1
End Sub


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
'   Decoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = IObject_Equals(Value)
End Function

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Bytes), Index, Count)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", Index, "Index", Count, "Count"
    
    Decoder_GetCharCount = InternalGetCharCount(Count)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount"
    
    Result = VerifyArrayRange(SAPtr(Chars), CharIndex, InternalGetCharCount(ByteCount))
    Select Case Result
        Case Argument_InvalidOffLen:   Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
        Case Is <> NO_ERROR:                ThrowArrayRangeException Result, "Chars", CharIndex, "CharIndex", 0, vbNullString
    End Select
    
    If mIsBigEndian Then
        Dim TotalBytes As Long
        If mLastByte >= 0 Then
            TotalBytes = ByteCount + 1
        Else
            TotalBytes = ByteCount
        End If
        
        Dim ch As Long
        Dim i As Long
        Do While i < TotalBytes
            If mLastByte >= 0 Then
                ch = mLastByte * &H100
                i = i - 1
                mLastByte = -1
            Else
                ch = Bytes(ByteIndex + i) * &H100
            End If
            i = i + 1
            ch = ch Or Bytes(ByteIndex + i)
            Chars(CharIndex) = AsWord(ch)
            CharIndex = CharIndex + 1
            i = i + 1
        Loop
        If (TotalBytes And 1) = 1 Then mLastByte = Bytes(ByteIndex + ByteCount - 1)
        Decoder_GetChars = TotalBytes \ 2
    Else
        Dim Offset As Long
        If mLastByte >= 0 Then
            Chars(CharIndex) = mLastByte
            Offset = 1
        End If
        Call CopyMemory(ByVal VarPtr(Chars(CharIndex)) + Offset, Bytes(ByteIndex), ByteCount - Offset)
        mLastByte = -1
        If ((ByteCount + Offset) And 1) = 1 Then mLastByte = Bytes(ByteIndex + ByteCount - 1)
        Decoder_GetChars = (ByteCount + Offset) \ 2
    End If
End Function

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = IObject_GetHashCode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = IObject_ToString
End Function

