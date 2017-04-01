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
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: UnicodeDecoder
'
Option Explicit
Implements IObject
Implements Decoder

'Private Const CHAR_SIZE As Long = 2

Private mEncoding       As UnicodeEncoding
Private mLastByte       As Long
Private mLastChar       As Long


Friend Property Get LastByte() As Long
    LastByte = mLastByte
End Property

Friend Property Let LastByte(ByVal Value As Long)
    mLastByte = Value
End Property

Friend Property Get LastChar() As Long
    LastChar = mLastChar
End Property

Friend Property Let LastChar(ByVal Value As Long)
    mLastChar = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mLastByte = -1
End Sub

Friend Sub Init(ByVal Encoding As UnicodeEncoding)
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
'   Decoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = IObject_Equals(Value)
End Function

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, Optional ByVal Flush As Boolean) As Long
    If SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    If Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - Index + 1 < Count Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
            
    Decoder_GetCharCount = mEncoding.GetCharCountCore(Bytes, Index, Count, Me)
    
    ' need look to deeper
    If Flush Then
        mLastByte = -1
        mLastChar = 0
    End If
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
'    Dim Result As Long
'    Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
'    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount"
'
'    Result = VerifyArrayRange(SAPtr(Chars), CharIndex, InternalGetCharCount(ByteCount))
'    Select Case Result
'        Case Argument_InvalidOffLen:   Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
'        Case Is <> NO_ERROR:                ThrowArrayRangeException Result, "Chars", CharIndex, "CharIndex", 0, vbNullString
'    End Select
'
'    If mIsBigEndian Then
'        Dim TotalBytes As Long
'        If mLastByte >= 0 Then
'            TotalBytes = ByteCount + 1
'        Else
'            TotalBytes = ByteCount
'        End If
'
'        Dim Ch As Long
'        Dim i As Long
'        Do While i < TotalBytes
'            If mLastByte >= 0 Then
'                Ch = mLastByte * &H100
'                i = i - 1
'                mLastByte = -1
'            Else
'                Ch = Bytes(ByteIndex + i) * &H100
'            End If
'            i = i + 1
'            Ch = Ch Or Bytes(ByteIndex + i)
'            Chars(CharIndex) = AsWord(Ch)
'            CharIndex = CharIndex + 1
'            i = i + 1
'        Loop
'        If (TotalBytes And 1) = 1 Then mLastByte = Bytes(ByteIndex + ByteCount - 1)
'        Decoder_GetChars = TotalBytes \ 2
'    Else
'        Dim Offset As Long
'        If mLastByte >= 0 Then
'            Chars(CharIndex) = mLastByte
'            Offset = 1
'        End If
'        Call CopyMemory(ByVal VarPtr(Chars(CharIndex)) + Offset, Bytes(ByteIndex), ByteCount - Offset)
'        mLastByte = -1
'        If ((ByteCount + Offset) And 1) = 1 Then mLastByte = Bytes(ByteIndex + ByteCount - 1)
'        Decoder_GetChars = (ByteCount + Offset) \ 2
'    End If
    
End Function

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = IObject_GetHashCode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = IObject_ToString
End Function

