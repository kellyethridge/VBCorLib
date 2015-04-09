Attribute VB_Name = "Comparisons"
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
' Module: Comparisons
'
Option Explicit

' This is a set of comparison routines used by function delegation calls.
' They allow a virtual comparison routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
Public Function SZCompareLongs(ByRef x As Long, ByRef y As Long) As Long
    If x > y Then
        SZCompareLongs = 1
    ElseIf x < y Then
        SZCompareLongs = -1
    End If
End Function

Public Function SZCompareIntegers(ByRef x As Integer, ByRef y As Integer) As Long
    If x > y Then
        SZCompareIntegers = 1
    ElseIf x < y Then
        SZCompareIntegers = -1
    End If
End Function

Public Function SZCompareStrings(ByRef x As String, ByRef y As String) As Long
    If x > y Then
        SZCompareStrings = 1
    ElseIf x < y Then
        SZCompareStrings = -1
    End If
End Function

Public Function SZCompareDoubles(ByRef x As Double, ByRef y As Double) As Long
    If x > y Then
        SZCompareDoubles = 1
    ElseIf x < y Then
        SZCompareDoubles = -1
    End If
End Function

Public Function SZCompareSingles(ByRef x As Single, ByRef y As Single) As Long
    If x > y Then
        SZCompareSingles = 1
    ElseIf x < y Then
        SZCompareSingles = -1
    End If
End Function

Public Function SZCompareBytes(ByRef x As Byte, ByRef y As Byte) As Long
    If x > y Then
        SZCompareBytes = 1
    ElseIf x < y Then
        SZCompareBytes = -1
    End If
End Function

Public Function SZCompareBooleans(ByRef x As Boolean, ByRef y As Boolean) As Long
    If x > y Then
        SZCompareBooleans = 1
    ElseIf x < y Then
        SZCompareBooleans = -1
    End If
End Function

Public Function SZCompareDates(ByRef x As Date, ByRef y As Date) As Long
    SZCompareDates = DateDiff("s", y, x)
End Function

Public Function SZCompareCurrencies(ByRef x As Currency, ByRef y As Currency) As Long
    If x > y Then
        SZCompareCurrencies = 1
    ElseIf x < y Then
        SZCompareCurrencies = -1
    End If
End Function

Public Function SZCompareComparables(ByRef x As Object, ByRef y As Variant) As Long
    Dim XComparable As IComparable
    Set XComparable = x
    SZCompareComparables = XComparable.CompareTo(y)
End Function

Public Function SZCompareVariants(ByRef x As Variant, ByRef y As Variant) As Long
    Dim Comparable As IComparable
    Dim XVarType As VbVarType
    Dim YVarType As VbVarType
    
    XVarType = VarType(x)
    YVarType = VarType(y)
    
    Select Case XVarType
        Case vbNull
            If YVarType <> vbNull Then
                SZCompareVariants = -1
            End If
            Exit Function

        Case vbEmpty
            If YVarType = vbNull Then
                SZCompareVariants = 1
            ElseIf YVarType <> vbEmpty Then
                SZCompareVariants = -1
            End If
            Exit Function
            
        Case vbObject, vbDataObject
            If Not x Is Nothing Then
                If TypeOf x Is IComparable Then
                    Set Comparable = x
                    SZCompareVariants = Comparable.CompareTo(y)
                    Exit Function
                End If
            End If
            
        Case YVarType
            If x < y Then
                SZCompareVariants = -1
            ElseIf x > y Then
                SZCompareVariants = 1
            End If
            Exit Function
    End Select

    Select Case YVarType
        Case vbNull, vbEmpty
            SZCompareVariants = 1
            
        Case vbObject, vbDataObject
            If TypeOf y Is IComparable Then
                Set Comparable = y
                SZCompareVariants = -Comparable.CompareTo(x)
            Else
                Throw Cor.NewArgumentException("Object must implement IComparable interface.")
            End If
            
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidValueType, TypeName(x)))
    End Select
End Function

' This is a set of equality routines used by function delegation calls.
' They allow a virtual equality routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
'
Public Function EqualsLong(ByRef x As Long, ByRef y As Long) As Boolean
    EqualsLong = (x = y)
End Function

Public Function EqualsString(ByRef x As String, ByRef y As String) As Boolean
    EqualsString = (x = y)
End Function

Public Function EqualsDouble(ByRef x As Double, ByRef y As Double) As Boolean
    EqualsDouble = (x = y)
End Function

Public Function EqualsInteger(ByRef x As Integer, ByRef y As Integer) As Boolean
    EqualsInteger = (x = y)
End Function

Public Function EqualsSingle(ByRef x As Single, ByRef y As Single) As Boolean
    EqualsSingle = (x = y)
End Function

Public Function EqualsDate(ByRef x As Date, ByRef y As Date) As Boolean
    EqualsDate = (DateDiff("s", x, y) = 0)
End Function

Public Function EqualsByte(ByRef x As Byte, ByRef y As Byte) As Boolean
    EqualsByte = (x = y)
End Function

Public Function EqualsBoolean(ByRef x As Boolean, ByRef y As Boolean) As Boolean
    EqualsBoolean = (x = y)
End Function

Public Function EqualsCurrency(ByRef x As Currency, ByRef y As Currency) As Boolean
    EqualsCurrency = (x = y)
End Function

Public Function EqualsObject(ByRef x As Object, ByRef y As Object) As Boolean
    If Not x Is Nothing Then
        If TypeOf x Is IObject Then
            Dim Obj As IObject
            Set Obj = x
            EqualsObject = Obj.Equals(y)
        Else
            EqualsObject = x Is y
        End If
    Else
        EqualsObject = y Is Nothing
    End If
End Function

Public Function EqualsVariants(ByRef x As Variant, ByRef y As Variant) As Boolean
    Dim o As IObject
    Select Case VarType(x)
        Case vbObject
            If x Is Nothing Then
                If IsObject(y) Then
                    EqualsVariants = (y Is Nothing)
                End If
            ElseIf TypeOf x Is IObject Then
                Set o = x
                EqualsVariants = o.Equals(y)
            ElseIf IsObject(y) Then
                If y Is Nothing Then Exit Function
                If TypeOf y Is IObject Then
                    Set o = y
                    EqualsVariants = o.Equals(x)
                Else
                    EqualsVariants = (x Is y)
                End If
            End If
        Case vbNull
            EqualsVariants = IsNull(y)
        Case VarType(y)
            EqualsVariants = (x = y)
    End Select
End Function
