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

Public Sub CheckValidSingleDimArray(ByRef Arr As Variant, Optional ByVal ParameterKey As ResourceStringKey = Parameter_Arr)
    Dim Ptr As Long
    Ptr = CorArray.ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr, ParameterKey
End Sub

Public Sub CheckValidSingleDimArrayPtr(ByVal Ptr As Long, Optional ByVal ParameterKey As ResourceStringKey = Parameter_Arr)
    If Ptr = vbNullPtr Then
        Error.ArgumentNull Environment.GetResourceString(ParameterKey), ArgumentNull_Array
    End If
    If SafeArrayGetDim(Ptr) <> 1 Then
        Error.Rank
    End If
End Sub

Public Sub ValidateArray(ByRef Arr As Variant, Optional ByVal ArrParameter As ParameterResourceKey = Parameter_Arr)
    Dim Ptr As Long
    Ptr = CorArray.ArrayPointer(Arr)
    If Ptr = vbNullPtr Then _
        Error.ArgumentNull Environment.GetParameterName(ArrParameter), ArgumentNull_Array
    If SafeArrayGetDim(Ptr) <> 1 Then _
        Error.Rank
End Sub

Public Sub ValidateArrayRange(ByRef Arr As Variant, ByVal Index As Long, ByVal Count As Long, Optional ByVal ArrParameter As ParameterResourceKey = Parameter_Arr, _
                                                                                              Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                              Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count)
    ValidateArray Arr, ArrParameter
    If Index < LBound(Arr) Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    If Count < 0 Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    If Index + Count - 1 > UBound(Arr) Then _
        Error.Argument Argument_InvalidOffLen
End Sub

Public Function ValidateArrayOptionalRange(ByRef Arr As Variant, ByRef Index As Variant, ByRef Count As Variant, Optional ByVal ArrParameter As ParameterResourceKey = Parameter_Arr, _
                                                                                                                 Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                                 Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count) As ListRange
    ValidateArray Arr, ArrParameter
    ValidateArrayOptionalRange = OptionalRange(Index, Count, LBound(Arr), CorArray.LengthFirstDim(Arr), IndexParameter, CountParameter)
    If ValidateArrayOptionalRange.Index < LBound(Arr) Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    If ValidateArrayOptionalRange.Count < 0 Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    If ValidateArrayOptionalRange.Index + ValidateArrayOptionalRange.Count - 1 > UBound(Arr) Then _
        Error.Argument Argument_InvalidOffLen
End Function

Public Sub ValidateByteArray(ByRef Bytes() As Byte, Optional ByVal Parameter As ParameterResourceKey = Parameter_Bytes)
    Dim Ptr As Long
    Ptr = SAPtr(Bytes)
    If Ptr = vbNullPtr Then _
        Error.ArgumentNull Environment.GetParameterName(Parameter), ArgumentNull_Array
    If SafeArrayGetDim(Ptr) <> 1 Then _
        Error.Rank
End Sub

Public Function ValidateByteArrayOptionalRange(ByRef Bytes() As Byte, ByRef Index As Variant, ByRef Count As Variant, Optional ByVal BytesParameter As ParameterResourceKey = Parameter_Bytes, _
                                                                                                                      Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                                      Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count) As ListRange
    ValidateByteArray Bytes, BytesParameter
    ValidateByteArrayOptionalRange = OptionalRange(Index, Count, LBound(Bytes), CorArray.LengthFirstDim(Bytes), IndexParameter, CountParameter)
    If ValidateByteArrayOptionalRange.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    If ValidateByteArrayOptionalRange.Count < 0 Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    If ValidateByteArrayOptionalRange.Index + ValidateByteArrayOptionalRange.Count - 1 > UBound(Bytes) Then _
        Error.Argument Argument_InvalidOffLen
End Function

Public Sub ValidateCharArray(ByRef Chars() As Integer, Optional ByVal Parameter As ParameterResourceKey = Parameter_Chars)
    Dim Ptr As Long
    Ptr = SAPtr(Chars)
    If Ptr = vbNullPtr Then _
        Error.ArgumentNull Environment.GetParameterName(Parameter), ArgumentNull_Array
    If SafeArrayGetDim(Ptr) <> 1 Then _
        Error.Rank
End Sub

Public Function ValidateCharArrayOptionalRange(ByRef Chars() As Integer, ByRef Index As Variant, ByRef Count As Variant, Optional ByVal CharsParameter As ParameterResourceKey = Parameter_Chars, _
                                                                                                                         Optional ByVal IndexParameter As ParameterResourceKey = Parameter_Index, _
                                                                                                                         Optional ByVal CountParameter As ParameterResourceKey = Parameter_Count) As ListRange
    ValidateCharArray Chars, CharsParameter
    ValidateCharArrayOptionalRange = OptionalRange(Index, Count, LBound(Chars), CorArray.LengthFirstDim(Chars), IndexParameter, CountParameter)
    If ValidateCharArrayOptionalRange.Index < LBound(Chars) Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(IndexParameter), ArgumentOutOfRange_LBound
    If ValidateCharArrayOptionalRange.Count < 0 Then _
        Error.ArgumentOutOfRange Environment.GetParameterName(CountParameter), ArgumentOutOfRange_NeedNonNegNum
    If ValidateCharArrayOptionalRange.Index + ValidateCharArrayOptionalRange.Count - 1 > UBound(Chars) Then _
        Error.Argument Argument_InvalidOffLen
End Function

