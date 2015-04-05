Attribute VB_Name = "Validate"
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


Public Sub ArrayRange(ByRef Arr As Variant, ByRef Range As ListRange, _
                              Optional ByVal IndexParameter As Param = Param.Index, _
                              Optional ByVal CountParameter As Param = Param.Count)
    If Range.Index < LBound(Arr) Then
        ThrowHelper.CannotBeLessThanLBound IndexParameter
    End If
    If Range.Count < 0 Then
        ThrowHelper.CannotBeNegative CountParameter
    End If
    If Range.Index + Range.Count - 1 > UBound(Arr) Then
        ThrowHelper.PositionNotValidForCollection
    End If
End Sub
