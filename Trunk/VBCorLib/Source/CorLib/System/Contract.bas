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

Public Sub CheckNull(ByVal ObjectToCheck As Object, ByVal Parameter As ParameterResourceId, ByVal Message As ResourceStringId)
    If ObjectToCheck Is Nothing Then
        Throw Cor.NewArgumentNullException(Resources.GetString(Parameter), Resources.GetString(Message))
    End If
End Sub

Public Sub CheckRange(ByVal Condition As Boolean, ByVal Parameter As ParameterResourceId, ByVal Message As ResourceStringId)
    If Not Condition Then
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(Parameter), Message:=Resources.GetString(Message))
    End If
End Sub

Public Sub CheckEmpty(ByRef StringToCheck As String, ByVal Parameter As ParameterResourceId, ByVal Message As ResourceStringId)
    If LenB(StringToCheck) = 0 Then
        Throw Cor.NewArgumentException(Resources.GetString(Message), Resources.GetString(Parameter))
    End If
End Sub

Public Sub CheckChars(ByRef Chars() As Integer)
    Dim CharPtr As Long
    CharPtr = SAPtr(Chars)
    If CharPtr = vbNullPtr Then
        Throw Cor.NewArgumentNullException(Resources.GetString(Param_Chars), Resources.GetString(ArgumentNull_Array))
    End If
    If SafeArrayGetDim(CharPtr) > 1 Then
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
    End If
End Sub

Public Sub CheckCharRange(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, Optional ByVal IndexParameter As ParameterResourceId = Param_Index, Optional ByVal CountParameter As ParameterResourceId = Param_Count)
    CheckRange Index >= LBound(Chars), IndexParameter, ArgumentOutOfRange_NeedNonNegNum
    CheckRange Count >= 0, CountParameter, ArgumentOutOfRange_NeedNonNegNum
    CheckRange Index + Count <= (UBound(Chars) - LBound(Chars) + 1), Param_Chars, ArgumentOutOfRange_IndexLength
End Sub
