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

Public Sub That(ByVal Assertion As Boolean, Optional ByVal Message As ResourceString = Argument_Exception, Optional ByVal ParameterName As ResourceString = Parameter_None)
    If Not Assertion Then
        Throw Cor.NewArgumentException(Environment.GetResourceString(Message), Environment.GetResourceString(ParameterName))
    End If
End Sub

Public Sub Range(ByVal Assertion As Boolean, Optional ByVal ParameterName As ResourceString = Parameter_None, Optional ByVal Message As ResourceString = ArgumentOutOfRange_Exception)
    If Not Assertion Then
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ParameterName), Message:=Environment.GetResourceString(Message))
    End If
End Sub

Public Sub NotNothing(ByVal Object As Object, Optional ByVal ParameterName As ResourceString = Parameter_None, Optional ByVal Message As ResourceString = ArgumentNull_Exception)
    If Object Is Nothing Then
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ParameterName), Environment.GetResourceString(Message))
    End If
End Sub

Public Sub NotNull(ByRef Arr As Variant, Optional ByVal ParameterName As ResourceString = Parameter_Arr, Optional ByVal Message As ResourceString = ArgumentNull_Array)
    Require.NotNullPtr ArrayPointer(Arr), ParameterName, Message
End Sub

Public Sub OneDimensionArray(ByRef Arr As Variant, Optional ByVal ParameterName As ResourceString = Parameter_Arr)
    Require.OneDimensionArrayPtr ArrayPointer(Arr), ParameterName
End Sub

Public Sub NotNullOneDimensionArray(ByRef Arr As Variant, Optional ByVal ParameterName As ResourceString = Parameter_Arr, Optional ByVal Message As ResourceString = ArgumentNull_Array)
    Dim ArrayPtr As Long
    ArrayPtr = ArrayPointer(Arr)
    
    Require.NotNullPtr ArrayPtr, ParameterName, Message
    Require.OneDimensionArrayPtr ArrayPtr, ParameterName
End Sub

Public Sub NotNullPtr(ByVal ArrayPtr As Long, Optional ByVal ParameterName As ResourceString = Parameter_Arr, Optional ByVal Message As ResourceString = ArgumentNull_Array)
    If ArrayPtr = vbNullPtr Then
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ParameterName), Environment.GetResourceString(Message))
    End If
End Sub

Public Sub OneDimensionArrayPtr(ByVal ArrayPtr As Long, Optional ByVal ParameterName As ResourceString = Parameter_Arr)
    If SafeArrayGetDim(ArrayPtr) > 1 Then
        Throw Cor.NewArgumentException(Environment.GetResourceString(Rank_MultiDimNotSupported), Environment.GetResourceString(ParameterName))
    End If
End Sub

