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
Option Explicit

Public Sub That(ByVal Condition As Boolean, ByVal Message As Argument, Optional ByVal Parameter As Param = Param.None)
    If Not Condition Then
        Throw Cor.NewArgumentException(Resources.GetString(Message), Resources.GetParameter(Parameter))
    End If
End Sub

Public Sub ObjectNotNull(ByVal ObjectToCheck As Object, Optional ByVal Parameter As Param = Param.None, Optional ByVal Message As ArgumentNull = ArgumentNull.NullGeneric)
    If ObjectToCheck Is Nothing Then
        Throw Cor.NewArgumentNullException(Resources.GetParameter(Parameter), Resources.GetMessage(Message))
    End If
End Sub

Public Sub ArrayNotNull(ByRef Arr As Variant, Optional ByVal Parameter As Param = Param.Arr, Optional ByVal Message As ArgumentNull = ArgumentNull.NullArray)
    Dim ArrayPtr As Long
    ArrayPtr = GetArrayPointer(Arr)
    
    VerifyArrayNotNull ArrayPtr, Parameter, Message
End Sub

Public Sub ArrayIs1Dimension(ByRef Arr As Variant, Optional ByVal Parameter As Param = Param.Arr)
    Dim ArrayPtr As Long
    ArrayPtr = GetArrayPointer(Arr)
     
    VerifyArray1Dimension ArrayPtr, Parameter
End Sub

Public Sub NotNull1DimensionArray(ByRef Arr As Variant, Optional ByVal Parameter As Param = Param.Arr, Optional ByVal Message As ArgumentNull = ArgumentNull.NullArray)
    Dim ArrayPtr As Long
    ArrayPtr = GetArrayPointer(Arr)
    
    VerifyArrayNotNull ArrayPtr, Parameter, Message
    VerifyArray1Dimension ArrayPtr, Parameter
End Sub

Public Sub ThatArgument(ByVal Condition As Boolean, Optional ByVal Parameter As Param = Param.None, Optional ByVal Message As ArgumentOutOfRange = ArgumentOutOfRange.Exception)
    If Not Condition Then
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetParameter(Parameter), Message:=Resources.GetMessage(Message))
    End If
End Sub

Private Sub VerifyArrayNotNull(ByVal ArrayPtr As Long, ByVal Parameter As Param, ByVal Message As ArgumentNull)
    If ArrayPtr = vbNullPtr Then
        Throw Cor.NewArgumentNullException(Resources.GetParameter(Parameter), Resources.GetMessage(Message))
    End If
End Sub

Private Sub VerifyArray1Dimension(ByVal ArrayPtr As Long, ByVal Parameter)
    If SafeArrayGetDim(ArrayPtr) > 1 Then
        Throw Cor.NewRankException(Resources.GetMessage(Rank.MultiDimensionNotSupported))
    End If
End Sub

