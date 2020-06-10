Attribute VB_Name = "Validation"
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
' Module: Validation
'
Option Explicit

Public Sub ValidateArrayRange(ByRef Arr As Variant, ByRef Index As Variant, ByRef Count As Variant, _
                              Optional ByVal ArrName As ParameterName = NameOfArr, _
                              Optional ByVal IndexName As ParameterName = NameOfIndex, _
                              Optional ByVal CountName As ParameterName = NameOfCount)
    Dim RangeIndex      As Long
    Dim RangeCount      As Long
    Dim IndexIsMissing  As Boolean
    
    ValidateArray Arr, ArrName
    IndexIsMissing = IsMissing(Index)
    
    If IndexIsMissing <> IsMissing(Count) Then
        Error.Argument Argument_ParamRequired, Environment.GetParameterName(IIf(IndexIsMissing, IndexName, CountName))
    End If
    
    If IndexIsMissing Then
        Exit Sub
    End If
    
    RangeIndex = Index
    RangeCount = Count
    
    If RangeIndex < LBound(Arr) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexName), ArgumentOutOfRange_LBound
    End If
    
    If RangeCount < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountName), ArgumentOutOfRange_NeedNonNegNum
    End If
    
    If RangeIndex + RangeCount - 1 > UBound(Arr) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Sub

Public Sub ValidateArray(ByRef Arr As Variant, Optional ByVal ArrName As ParameterName = NameOfArr)
    Dim Ptr As Long
    
    If Not IsArray(Arr) Then
        Error.Argument Argument_ArrayRequired, Environment.GetParameterName(ArrName)
    End If
    
    Ptr = SAPtrV(Arr)
    
    If Ptr = vbNullPtr Then
        Error.ArgumentNull Environment.GetParameterName(ArrName), ArgumentNull_Array
    End If
    
    If SafeArrayGetDim(Ptr) <> 1 Then
        Error.Rank
    End If
End Sub

Public Sub ValidateByteRange(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, _
                             Optional ByVal BytesName As ParameterName = NameOfBytes, _
                             Optional ByVal IndexName As ParameterName = NameOfIndex, _
                             Optional ByVal CountName As ParameterName = NameOfCount)
    If SAPtr(Bytes) = vbNullPtr Then
        Error.ArgumentNull Environment.GetParameterName(BytesName), ArgumentNull_Array
    End If
    
    If Index < LBound(Bytes) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexName), ArgumentOutOfRange_LBound
    End If
    
    If Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountName), ArgumentOutOfRange_NeedNonNegNum
    End If
    
    If Index + Count - 1 > UBound(Bytes) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Sub

Public Sub ValidateCharRange(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, _
                             Optional ByVal CharsName As ParameterName = NameOfChars, _
                             Optional ByVal IndexName As ParameterName = NameOfIndex, _
                             Optional ByVal CountName As ParameterName = NameOfCount)
    If SAPtr(Chars) = vbNullPtr Then
        Error.ArgumentNull Environment.GetParameterName(CharsName), ArgumentNull_Array
    End If
    
    If Index < LBound(Chars) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexName), ArgumentOutOfRange_LBound
    End If
    
    If Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountName), ArgumentOutOfRange_NeedNonNegNum
    End If
    
    If Index + Count - 1 > UBound(Chars) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Sub


