Attribute VB_Name = "Require"
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
' Module: Require
'

''
' Provides argument validation methods.
'
Option Explicit

Public Sub That(ByVal Assertion As Boolean, Optional ByVal Message As ErrorMessage = Argument_Exception, Optional ByVal Parameter As ParameterName = Parameter_None)
    If Not Assertion Then
        Throw Cor.NewArgumentException(GetErrorMessage(Message), GetParameterName(Parameter))
    End If
End Sub

Public Sub Range(ByVal Assertion As Boolean, Optional ByVal Parameter As ParameterName = Parameter_None, Optional ByVal Message As ErrorMessage = ArgumentOutOfRange_Exception)
    If Not Assertion Then
        Throw Cor.NewArgumentOutOfRangeException(GetParameterName(Parameter), Message:=GetErrorMessage(Message))
    End If
End Sub

Public Sub ArrayRange(ByRef Arr As Variant, ByRef Range As ListRange, Optional ByVal IndexParameter As ParameterName = Parameter_Index, Optional ByVal CountParameter As ParameterName = Parameter_Count, Optional ByVal ArrayParameter As ParameterName = Parameter_Arr)
    Require.Range Range.Index >= LBound(Arr), IndexParameter, ArgumentOutOfRange_ArrayLB
    Require.Range Range.Count >= 0, CountParameter, ArgumentOutOfRange_NeedNonNegNum
    Require.Range Range.Index + Range.Count - 1 <= UBound(Arr), ArrayParameter, Argument_InvalidOffLen
End Sub

Public Sub NotNothing(ByVal Object As Object, Optional ByVal Parameter As ParameterName = Parameter_None, Optional ByVal Message As ErrorMessage = ArgumentNull_Exception)
    If Object Is Nothing Then
        Throw Cor.NewArgumentNullException(GetParameterName(Parameter), GetErrorMessage(Message))
    End If
End Sub

Public Sub NotNull(ByRef Arr As Variant, Optional ByVal Parameter As ParameterName = Parameter_Arr, Optional ByVal Message As ErrorMessage = ArgumentNull_Array)
    Require.NotNullPtr ArrayPointer(Arr), Parameter, Message
End Sub

Public Sub OneDimensionArray(ByRef Arr As Variant, Optional ByVal Parameter As ParameterName = Parameter_Arr)
    Require.OneDimensionArrayPtr ArrayPointer(Arr), Parameter
End Sub

Public Sub NotNullOneDimensionArray(ByRef Arr As Variant, Optional ByVal Parameter As ParameterName = Parameter_Arr, Optional ByVal Message As ErrorMessage = ArgumentNull_Array)
    Dim ArrayPtr As Long
    ArrayPtr = ArrayPointer(Arr)
    
    Require.NotNullPtr ArrayPtr, Parameter, Message
    Require.OneDimensionArrayPtr ArrayPtr, Parameter
End Sub

Public Sub NotNullPtr(ByVal ArrayPtr As Long, Optional ByVal Parameter As ParameterName = Parameter_Arr, Optional ByVal Message As ErrorMessage = ArgumentNull_Array)
    If ArrayPtr = vbNullPtr Then
        Throw Cor.NewArgumentNullException(GetParameterName(Parameter), GetErrorMessage(Message))
    End If
End Sub

Public Sub OneDimensionArrayPtr(ByVal ArrayPtr As Long, Optional ByVal Parameter As ParameterName = Parameter_Arr)
    If SafeArrayGetDim(ArrayPtr) > 1 Then
        Throw Cor.NewArgumentException(GetErrorMessage(Rank_MultiDimNotSupported), GetParameterName(Parameter))
    End If
End Sub

