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
Public Function CompareLongs(ByRef x As Long, ByRef y As Long) As Long
    If x > y Then
        CompareLongs = 1
    ElseIf x < y Then
        CompareLongs = -1
    End If
End Function
Public Function CompareIntegers(ByRef x As Integer, ByRef y As Integer) As Long
    If x > y Then
        CompareIntegers = 1
    ElseIf x < y Then
        CompareIntegers = -1
    End If
End Function
Public Function CompareStrings(ByRef x As String, ByRef y As String) As Long
    If x > y Then
        CompareStrings = 1
    ElseIf x < y Then
        CompareStrings = -1
    End If
End Function
Public Function CompareDoubles(ByRef x As Double, ByRef y As Double) As Long
    If x > y Then
        CompareDoubles = 1
    ElseIf x < y Then
        CompareDoubles = -1
    End If
End Function
Public Function CompareSingles(ByRef x As Single, ByRef y As Single) As Long
    If x > y Then
        CompareSingles = 1
    ElseIf x < y Then
        CompareSingles = -1
    End If
End Function
Public Function CompareBytes(ByRef x As Byte, ByRef y As Byte) As Long
    If x > y Then
        CompareBytes = 1
    ElseIf x < y Then
        CompareBytes = -1
    End If
End Function
Public Function CompareBooleans(ByRef x As Boolean, ByRef y As Boolean) As Long
    If x > y Then
        CompareBooleans = 1
    ElseIf x < y Then
        CompareBooleans = -1
    End If
End Function
Public Function CompareDates(ByRef x As Date, ByRef y As Date) As Long
    CompareDates = DateDiff("s", y, x)
End Function
Public Function CompareCurrencies(ByRef x As Currency, ByRef y As Currency) As Long
    If x > y Then CompareCurrencies = 1: Exit Function
    If x < y Then CompareCurrencies = -1
End Function
Public Function CompareIComparable(ByRef x As Object, ByRef y As Variant) As Long
    Dim comparableX As IComparable
    Set comparableX = x
    CompareIComparable = comparableX.CompareTo(y)
End Function

Public Function CompareVariants(ByRef x As Variant, ByRef y As Variant) As Long
    Dim Comparable As IComparable
    
    If VarType(x) <> VarType(y) Then _
        Throw Cor.NewArgumentException("A value of type " & TypeName(x) & " is required.")
    
    Select Case VarType(x)
        Case vbNull
            Exit Function
            
'        Case vbNull, vbEmpty
'        Case vbNull
'            If Not IsNull(y) Then
'                CompareVariants = -1
'            End If
'
'        Case vbEmpty
'            If IsNull(y) Then
'                CompareVariants = 1
'            ElseIf Not IsEmpty(y) Then
'                CompareVariants = -1
'            End If

        Case vbObject, vbDataObject
            If TypeOf x Is IComparable Then
                Set Comparable = x
                CompareVariants = Comparable.CompareTo(y)
            End If

        Case VarType(y)
            If x < y Then
                CompareVariants = -1
            ElseIf x > y Then
                CompareVariants = 1
            End If

    End Select
    
'    Select Case VarType(y)
'        Case vbNull, vbEmpty
'            CompareVariants = 1
'        Case vbObject, vbDataObject
'            If TypeOf y Is IComparable Then
'                Set Comparable = y
'                CompareVariants = -Comparable.CompareTo(x)
'                Exit Function
'            Else
'                Throw Cor.NewArgumentException("Object must implement IComparable interface.")
'            End If
'        Case Else
'            Throw Cor.NewInvalidOperationException("Specified IComparer failed.")
'    End Select
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
