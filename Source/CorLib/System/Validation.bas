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

Public Sub ValidateArray(ByRef Arr As Variant, Optional ByVal Parameter As ParameterResourceKey = Parameter_Arr)
    If Not IsArray(Arr) Then _
        Error.Argument Argument_ArrayRequired, Environment.GetParameterName(Parameter)
        
    ValidateArrayPtr VSAPtr(Arr), Parameter
End Sub

Public Sub ValidateArrayRange(ByRef Arr As Variant, ByVal Index As Long, ByVal Count As Long, Optional ByVal ArrParameter As ParameterResourceKey = Parameter_Arr, _
                                                                                              Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                              Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count)
    ValidateArray Arr, ArrParameter
    If Index < LBound(Arr) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    End If
    If Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    End If
    If Index + Count - 1 > UBound(Arr) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Sub

Public Function ValidateOptionalArrayRange(ByRef Arr As Variant, ByRef Index As Variant, ByRef Count As Variant, Optional ByVal ArrParameter As ParameterResourceKey = Parameter_Arr, _
                                                                                                                 Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                                 Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count) As ListRange
    ValidateArray Arr, ArrParameter
    ValidateOptionalArrayRange = GetOptionalRange(Index, Count, LBound(Arr), CorArray.LengthFirstDim(Arr), IndexParameter, CountParameter)
    If ValidateOptionalArrayRange.Index < LBound(Arr) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    End If
    If ValidateOptionalArrayRange.Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    End If
    If ValidateOptionalArrayRange.Index + ValidateOptionalArrayRange.Count - 1 > UBound(Arr) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Function

Public Sub ValidateByteArray(ByRef Bytes() As Byte, Optional ByVal Parameter As ParameterResourceKey = Parameter_Bytes)
    ValidateArrayPtr SAPtr(Bytes), Parameter
End Sub

Public Sub ValidateByteArrayRange(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, Optional ByVal BytesParameter As ParameterResourceKey = Parameter_Bytes, _
                                                                                                   Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                   Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count)
    ValidateByteArray Bytes, BytesParameter
    If Index < LBound(Bytes) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    End If
    If Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    End If
    If Index + Count - 1 > UBound(Bytes) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Sub

Public Function ValidateByteArrayOptionalRange(ByRef Bytes() As Byte, ByRef Index As Variant, ByRef Count As Variant, Optional ByVal BytesParameter As ParameterResourceKey = Parameter_Bytes, _
                                                                                                                      Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                                      Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count) As ListRange
    ValidateByteArray Bytes, BytesParameter
    ValidateByteArrayOptionalRange = GetOptionalRange(Index, Count, LBound(Bytes), CorArray.LengthFirstDim(Bytes), IndexParameter, CountParameter)
    If ValidateByteArrayOptionalRange.Index < LBound(Bytes) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    End If
    If ValidateByteArrayOptionalRange.Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    End If
    If ValidateByteArrayOptionalRange.Index + ValidateByteArrayOptionalRange.Count - 1 > UBound(Bytes) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Function

Public Sub ValidateCharArray(ByRef Chars() As Integer, Optional ByVal Parameter As ParameterResourceKey = Parameter_Chars)
    ValidateArrayPtr SAPtr(Chars), Parameter
End Sub

Public Sub ValidateCharArrayRange(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, Optional ByVal CharsParameter As ParameterResourceKey = Parameter_Chars, _
                                                                                                      Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                      Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count)
    ValidateCharArray Chars, CharsParameter
    If Index < LBound(Chars) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    End If
    If Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    End If
    If Index + Count - 1 > UBound(Chars) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Sub

Public Function ValidateCharArrayOptionalRange(ByRef Chars() As Integer, ByRef Index As Variant, ByRef Count As Variant, Optional ByVal CharsParameter As ParameterResourceKey = Parameter_Chars, _
                                                                                                                         Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                                         Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count) As ListRange
    ValidateCharArray Chars, CharsParameter
    ValidateCharArrayOptionalRange = GetOptionalRange(Index, Count, LBound(Chars), CorArray.LengthFirstDim(Chars), IndexParameter, CountParameter)
    If ValidateCharArrayOptionalRange.Index < LBound(Chars) Then
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    End If
    If ValidateCharArrayOptionalRange.Count < 0 Then
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    End If
    If ValidateCharArrayOptionalRange.Index + ValidateCharArrayOptionalRange.Count - 1 > UBound(Chars) Then
        Error.Argument Argument_InvalidOffLen
    End If
End Function

Private Sub ValidateArrayPtr(ByVal Ptr As Long, ByVal Parameter As ParameterResourceKey)
    If Ptr = vbNullPtr Then
        Error.ArgumentNull Environment.GetParameterName(Parameter), ArgumentNull_Array
    End If
    If SafeArrayGetDim(Ptr) <> 1 Then
        Error.Rank
    End If
End Sub

