Attribute VB_Name = "Contract"
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: Contract
'
Option Explicit

Public Sub Check(ByVal SuccessfulCondition As Boolean, ByVal Message As Argument, Optional ByVal Parameter As Param = Param.None)
    If Not SuccessfulCondition Then
        Dim ParameterName As String
        If Parameter <> Param.None Then
            ParameterName = Resources.GetParameter(Parameter)
        End If
        
        Throw Cor.NewArgumentException(Resources.GetString(Message), ParameterName)
    End If
End Sub

Public Sub CheckNot(ByVal FailingCondition As Boolean, ByVal Message As Argument, Optional ByVal Parameter As Param = Param.None)
    Check Not FailingCondition, Message, Parameter
End Sub

Public Sub CheckInRange(ByVal SuccessfulCondition As Boolean, ByVal Parameter As Param, ByVal Message As ArgumentOutOfRange)
    If Not SuccessfulCondition Then
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(Parameter), Message:=Resources.GetString(Message))
    End If
End Sub

Public Sub CheckNotInRange(ByVal FailingCondition As Boolean, ByVal Parameter As Param, ByVal Message As ArgumentOutOfRange)
    CheckInRange Not FailingCondition, Parameter, Message
End Sub

Public Sub CheckNotNull(ByVal ObjectToCheck As Object, ByVal Parameter As Param, Optional ByVal Message As ErrorMessage = ArgumentNull_Generic)
    If ObjectToCheck Is Nothing Then
        Throw Cor.NewArgumentNullException(Resources.GetString(Parameter), Resources.GetString(Message))
    End If
End Sub

Public Sub CheckArgument(ByVal FailingCondition As Boolean, ByVal Message As Argument, Optional ByVal Parameter As Param = Param.None)
    If FailingCondition Then
        Dim ParameterName As String
        If Parameter <> Param.None Then
            ParameterName = Resources.GetParameter(Parameter)
        End If
        
        Throw Cor.NewArgumentException(Resources.GetString(Message), ParameterName)
    End If
End Sub

Public Sub CheckNull(ByVal ValueToCheck As Object, ByVal Parameter As Param, Optional ByVal Message As ErrorMessage = ArgumentNull_Generic)
    If ValueToCheck Is Nothing Then
        Throw Cor.NewArgumentNullException(Resources.GetString(Parameter), Resources.GetString(Message))
    End If
End Sub

Public Sub CheckRange(ByVal FailingCondition As Boolean, ByVal Parameter As Param, ByVal Message As ArgumentOutOfRange)
    If FailingCondition Then
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(Parameter), Message:=Resources.GetString(Message))
    End If
End Sub

Public Sub CheckEmpty(ByRef StringToCheck As String, ByVal Parameter As Param, ByVal Message As ErrorMessage)
    If LenB(StringToCheck) = 0 Then
        Throw Cor.NewArgumentException(Resources.GetString(Message), Resources.GetString(Parameter))
    End If
End Sub

Public Sub CheckArray(ByRef ArrayToCheck As Variant, Optional Parameter As Param = Param.Bytes)
    Dim Ptr As Long
    Ptr = GetArrayPointer(ArrayToCheck)
    
    If Ptr = vbNullPtr Then
        Throw Cor.NewArgumentNullException(Resources.GetString(Parameter), Resources.GetString(ArgumentNull_Array))
    End If
    If SafeArrayGetDim(Ptr) > 1 Then
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
    End If
End Sub

Public Sub CheckArrayRange(ByRef ArrayToCheck As Variant, ByVal Index As Long, ByVal Count As Long, Optional ByVal IndexParameter As Param = Param.Index, Optional ByVal CountParameter As Param = Param.Count, Optional ByVal ArrayParameter As Param = Param.Bytes)
    CheckRange Index < LBound(ArrayToCheck), IndexParameter, ArgumentOutOfRange.LowerBound
    CheckRange Count < 0, CountParameter, ArgumentOutOfRange.NeedNonNegNum
    CheckRange Index + Count > UBound(ArrayToCheck) + 1, ArrayParameter, ArgumentOutOfRange.IndexLength
End Sub

Public Sub CheckArrayIndex(ByRef ArrayToCheck As Variant, ByVal Index As Long, Optional ByVal IndexParameter As Param = Param.Index)
    CheckRange Index < LBound(ArrayToCheck), IndexParameter, ArgumentOutOfRange.LowerBound
    CheckRange Index > UBound(ArrayToCheck), IndexParameter, ArgumentOutOfRange.UpperBound
End Sub

Private Sub VerifyArrayNotNull(ByVal ArrayPtr As Long, ByVal Parameter As Param, ByVal Message As ArgumentNull)
    If ArrayPtr = vbNullPtr Then
        Throw Cor.NewArgumentNullException(Resources.GetParameter(Parameter), Resources.GetErrorMessage(Message))
    End If
End Sub

Private Sub VerifyArrayOneDimension(ByVal ArrayPtr As Long, ByVal Parameter As Param)
    If SafeArrayGetDim(ArrayPtr) > 1 Then
        Throw Cor.NewArgumentException(Resources.GetErrorMessage(Rank_MultiDimNotSupported), Resources.GetParameter(Parameter))
    End If
End Sub








