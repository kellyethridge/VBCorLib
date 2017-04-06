Attribute VB_Name = "modEncoding"
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
' Module: modEncoding
'
Option Explicit

Private Const Base64Characters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

' Cache the Base64 encoded character lookup table for quick access.
Public Base64CharToBits()   As Long
Public Base64Bytes()        As Byte


' Initialize the encoded character lookup table.
Public Sub InitEncoding()
    Dim i As Long
    
    ReDim Base64CharToBits(0 To 127)
    For i = 0 To 127
        Base64CharToBits(i) = vbInvalidChar
    Next i
    
    For i = 0 To 25
        Base64CharToBits(vbUpperAChar + i) = i
        Base64CharToBits(vbLowerAChar + i) = i + 26
    Next i
    
    For i = 0 To 9
        Base64CharToBits(vbZeroChar + i) = i + 52
    Next i
    
    Base64CharToBits(43) = vbGreaterThanChar
    Base64CharToBits(47) = vbQuestionMarkChar
    
    ReDim Base64Bytes(0 To 63)
    For i = 0 To Len(Base64Characters) - 1
        Base64Bytes(i) = Asc(Mid$(Base64Characters, i + 1, 1))
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Common methods shared by Encoding implementations
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DecoderConvert(ByVal Decoder As Decoder, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByVal Flush As Boolean, ByRef BytesUsed As Long, ByRef CharsUsed As Long, ByRef Completed As Boolean)
    If SAPtr(Bytes) = vbNullPtr Or SAPtr(Chars) = vbNullPtr Then _
        Error.ArgumentNull IIf(SAPtr(Bytes) = vbNullPtr, "Bytes", "Chars"), ArgumentNull_Array
    If ByteIndex < LBound(Bytes) Or CharIndex < LBound(Chars) Then _
        Error.ArgumentOutOfRange IIf(ByteIndex < LBound(Bytes), "ByteIndex", "CharIndex"), ArgumentOutOfRange_ArrayLB
    If ByteCount < 0 Or CharCount < 0 Then _
        Error.ArgumentOutOfRange IIf(ByteCount < 0, "ByteCount", "CharCount"), ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - ByteIndex + 1 < ByteCount Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    If UBound(Chars) - CharIndex + 1 < CharCount Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
    
    BytesUsed = ByteCount
    
    Do While BytesUsed > 0
        If Decoder.GetCharCount(Bytes, ByteIndex, BytesUsed, Flush) <= CharCount Then
            CharsUsed = Decoder.GetChars(Bytes, ByteIndex, BytesUsed, Chars, CharIndex, Flush)
            Completed = (BytesUsed = ByteCount) And (Decoder.FallbackBuffer.Remaining = 0)
            Exit Sub
        End If
        
        Flush = False
        BytesUsed = BytesUsed \ 2
    Loop
    
    Error.Argument Argument_ConversionOverflow
End Sub

Public Sub EncoderConvert(ByVal Encoder As Encoder, Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByVal Flush As Boolean, CharsUsed As Long, BytesUsed As Long, Completed As Boolean)
    If SAPtr(Chars) = vbNullPtr Or SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull IIf(SAPtr(Chars) = vbNullPtr, "Chars", "Bytes"), ArgumentNull_Array
    If CharIndex < LBound(Chars) Or ByteIndex < LBound(Bytes) Then _
        Error.ArgumentOutOfRange IIf(CharIndex < LBound(Chars), "CharIndex", "ByteIndex"), ArgumentOutOfRange_ArrayLB
    If CharCount < 0 Or ByteCount < 0 Then _
        Error.ArgumentOutOfRange IIf(CharCount < 0, "CharCount", "ByteCount"), ArgumentOutOfRange_NeedNonNegNum
    If UBound(Chars) - CharIndex + 1 < CharCount Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
    If UBound(Bytes) - ByteIndex + 1 < ByteCount Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
        
    CharsUsed = CharCount
    
    Do While CharsUsed > 0
        If Encoder.GetByteCount(Chars, CharIndex, CharsUsed, Flush) <= ByteCount Then
            BytesUsed = Encoder.GetBytes(Chars, CharIndex, CharsUsed, Bytes, ByteIndex, Flush)
            Completed = (CharsUsed = CharCount) And (Encoder.FallbackBuffer.Remaining = 0)
            Exit Sub
        End If
        
        Flush = False
        CharsUsed = CharsUsed \ 2
    Loop
    
    Error.Argument Argument_ConversionOverflow
End Sub

Public Function ValidateGetBytes(ByRef Chars() As Integer, ByRef Index As Variant, ByRef Count As Variant) As ListRange
    If SAPtr(Chars) = vbNullPtr Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
    
    ValidateGetBytes = GetArrayRange(Chars, Index, Count)
    
    If ValidateGetBytes.Index < LBound(Chars) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If ValidateGetBytes.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Chars) - ValidateGetBytes.Index + 1 < ValidateGetBytes.Count Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
End Function

Public Sub ValidateGetBytesEx(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long)
    If SAPtr(Chars) = vbNullPtr Or SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull IIf(SAPtr(Chars) = vbNullPtr, "Chars", "Bytes"), ArgumentNull_Array
    If CharIndex < LBound(Chars) Then _
        Error.ArgumentOutOfRange "CharIndex", ArgumentOutOfRange_ArrayLB
    If ByteIndex < LBound(Bytes) Or (ByteIndex > UBound(Bytes) And Len1D(Bytes) > 0) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_Index
    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Chars) - CharIndex + 1 < CharCount Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
End Sub

Public Function ValidateGetChars(ByRef Bytes() As Byte, ByRef Index As Variant, ByRef Count As Variant) As ListRange
    If SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    
    ValidateGetChars = GetArrayRange(Bytes, Index, Count)
    
    If ValidateGetChars.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If ValidateGetChars.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - ValidateGetChars.Index + 1 < ValidateGetChars.Count Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
End Function

Public Sub ValidateGetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long)
    If SAPtr(Bytes) = vbNullPtr Or SAPtr(Chars) = vbNullPtr Then _
        Error.ArgumentNull IIf(SAPtr(Bytes) = vbNullPtr, "Bytes", "Chars"), ArgumentNull_Array
    If ByteIndex < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_ArrayLB
    If CharIndex < LBound(Chars) Or (CharIndex > UBound(Chars) And Len1D(Chars) > 0) Then _
        Error.ArgumentOutOfRange "CharIndex", ArgumentOutOfRange_Index
    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - ByteIndex + 1 < ByteCount Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
End Sub

''
' Attaches either an Integer Array or a String to a Chars Integer
' array, allowing the same access type to both source types.
'
' @param Source Either an Integer Array or a String to attach to.
' @param Chars The array that will be used to access the elements in Source.
' @param CharsSA The SafeArray structure used to represent Chars.
'
Public Function AttachChars(ByRef Source As Variant, ByRef Chars() As Integer, ByRef CharsSA As SafeArray1d) As Long
    Select Case VarType(Source)
        Case vbString
            CharsSA.cElements = Len(Source)
            CharsSA.pvData = StrPtr(Source)
            CharsSA.cbElements = 2
            CharsSA.cDims = 1
            SAPtr(Chars) = VarPtr(CharsSA)
            AttachChars = Len(Source)
            
        Case vbIntegerArray
            Dim CharPtr As Long
            CharPtr = CorArray.ArrayPointer(Source)
            If CharPtr = vbNullPtr Then _
                Throw Cor.NewArgumentNullException(Environment.GetResourceString(Parameter_Chars), Environment.GetResourceString(ArgumentNull_Array))
            If SafeArrayGetDim(CharPtr) > 1 Then _
                Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimNotSupported))
            
            SAPtr(Chars) = CharPtr
            AttachChars = UBound(Chars) - LBound(Chars) + 1
            
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_CharArrayRequired), Environment.GetResourceString(Parameter_Chars))
    End Select
End Function

