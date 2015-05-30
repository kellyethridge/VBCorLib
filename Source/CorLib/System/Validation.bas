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
        Throw Error.ArgumentNull(Environment.GetResourceString(ParameterKey), ArgumentNull_Array)
    End If
    If SafeArrayGetDim(Ptr) <> 1 Then
        Throw Error.Rank
    End If
End Sub

