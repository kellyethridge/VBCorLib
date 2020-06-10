Attribute VB_Name = "Interaction"
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: CorInteraction
'
Option Explicit

Public Function IfObject(ByVal ObjA As Object, ByVal ObjB As Object) As Object
    If ObjA Is Nothing Then
        Set IfObject = ObjB
    Else
        Set IfObject = ObjA
    End If
End Function

Public Function IfString(ByRef a As String, ByRef b As String) As String
    If LenB(a) > 0 Then
        IfString = a
    Else
        IfString = b
    End If
End Function

Public Function IIfLong(ByVal Expression As Boolean, ByVal TruePart As Long, ByVal FalsePart As Long) As Long
    If Expression Then
        IIfLong = TruePart
    Else
        IIfLong = FalsePart
    End If
End Function

Public Sub SwapByte(ByRef a As Byte, ByRef b As Byte)
    Dim t As Byte
    t = a: a = b: b = t
End Sub

Public Sub SwapInteger(ByRef a As Integer, ByRef b As Integer)
    Dim t As Integer
    t = a: a = b: b = t
End Sub

Public Sub SwapLong(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a: a = b: b = t
End Sub

Public Sub SwapCurrency(ByRef a As Currency, ByRef b As Currency)
    Dim t As Currency
    t = a: a = b: b = t
End Sub

Public Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
    Helper.Swap16 a, b
End Sub
