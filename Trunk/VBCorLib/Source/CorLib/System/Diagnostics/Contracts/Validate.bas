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
' Module: Validate
'
Option Explicit


Public Sub ValidateArray(ByRef Arr As Variant, Optional ByVal Parameter As ParameterString = Parameter_Arr)
    Dim ArrayPtr As Long
    ArrayPtr = ArrayPointer(Arr)
    
    If ArrayPtr = vbNullPtr Then
        Error.ArgumentNull Parameter, ArgumentNull_Array
    End If
    If SafeArrayGetDim(ArrayPtr) > 1 Then
        Error.Rank
    End If
End Sub

Public Sub ValidateArrayRange(ByRef Range As ListRange, ByRef Arr As Variant, Optional ByVal IndexParameter As ParameterString = Parameter_Index, Optional ByVal CountParameter As ParameterString = Parameter_Count)
    If Range.Index < LBound(Arr) Then
        Error.ArgumentOutOfRange IndexParameter, ArgumentOutOfRange_LBound
    End If
    If Range.Count < 0 Then
        Error.NegativeNumber CountParameter
    End If
    If Range.Index + Range.Count - 1 > UBound(Arr) Then
        Error.InvalidOffsetLength
    End If
End Sub

Public Sub ValidateListRange(ByRef Range As ListRange, ByVal ListCount As Long, Optional ByVal IndexParameter As ParameterString = Parameter_Index, Optional ByVal CountParameter As ParameterString = Parameter_Count)
    If Range.Index < 0 Then
        Error.NegativeNumber IndexParameter
    End If
    If Range.Count < 0 Then
        Error.NegativeNumber CountParameter
    End If
    If Range.Index + Range.Count > ListCount Then
        Error.InvalidOffsetLength
    End If
End Sub

