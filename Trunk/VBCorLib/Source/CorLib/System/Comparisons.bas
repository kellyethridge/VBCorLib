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


Public Function SZCompareVariants(ByRef X As Variant, ByRef Y As Variant) As Long
    Dim Comparable As IComparable
    Dim XVarType As VbVarType
    Dim YVarType As VbVarType
    
    XVarType = VarType(X)
    YVarType = VarType(Y)
    
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
            If Not X Is Nothing Then
                If TypeOf X Is IComparable Then
                    Set Comparable = X
                    SZCompareVariants = Comparable.CompareTo(Y)
                    Exit Function
                End If
            End If
            
        Case YVarType
            If X < Y Then
                SZCompareVariants = -1
            ElseIf X > Y Then
                SZCompareVariants = 1
            End If
            Exit Function
    End Select

    Select Case YVarType
        Case vbNull, vbEmpty
            SZCompareVariants = 1
            
        Case vbObject, vbDataObject
            If TypeOf Y Is IComparable Then
                Set Comparable = Y
                SZCompareVariants = -Comparable.CompareTo(X)
            Else
                Throw Cor.NewArgumentException("Object must implement IComparable interface.")
            End If
            
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidValueType, TypeName(X)))
    End Select
End Function

' This is a set of equality routines used by function delegation calls.
' They allow a virtual equality routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
'
Public Function EqualsLong(ByRef X As Long, ByRef Y As Long) As Boolean
    EqualsLong = (X = Y)
End Function

Public Function EqualsString(ByRef X As String, ByRef Y As String) As Boolean
    EqualsString = (X = Y)
End Function

Public Function EqualsDouble(ByRef X As Double, ByRef Y As Double) As Boolean
    EqualsDouble = (X = Y)
End Function

Public Function EqualsInteger(ByRef X As Integer, ByRef Y As Integer) As Boolean
    EqualsInteger = (X = Y)
End Function

Public Function EqualsSingle(ByRef X As Single, ByRef Y As Single) As Boolean
    EqualsSingle = (X = Y)
End Function

Public Function EqualsDate(ByRef X As Date, ByRef Y As Date) As Boolean
    EqualsDate = (DateDiff("s", X, Y) = 0)
End Function

Public Function EqualsByte(ByRef X As Byte, ByRef Y As Byte) As Boolean
    EqualsByte = (X = Y)
End Function

Public Function EqualsBoolean(ByRef X As Boolean, ByRef Y As Boolean) As Boolean
    EqualsBoolean = (X = Y)
End Function

Public Function EqualsCurrency(ByRef X As Currency, ByRef Y As Currency) As Boolean
    EqualsCurrency = (X = Y)
End Function

Public Function EqualsObject(ByRef X As Object, ByRef Y As Object) As Boolean
    If Not X Is Nothing Then
        If TypeOf X Is IObject Then
            Dim Obj As IObject
            Set Obj = X
            EqualsObject = Obj.Equals(Y)
        Else
            EqualsObject = X Is Y
        End If
    Else
        EqualsObject = Y Is Nothing
    End If
End Function

Public Function EqualsVariants(ByRef X As Variant, ByRef Y As Variant) As Boolean
    Dim Obj As IObject
    
    Select Case VarType(X)
        Case vbObject
            If X Is Nothing Then
                If IsObject(Y) Then
                    EqualsVariants = (Y Is Nothing)
                End If
            ElseIf TypeOf X Is IObject Then
                Set Obj = X
                EqualsVariants = Obj.Equals(Y)
            ElseIf IsObject(Y) Then
                If Y Is Nothing Then
                    Exit Function
                End If
                
                If TypeOf Y Is IObject Then
                    Set Obj = Y
                    EqualsVariants = Obj.Equals(X)
                Else
                    EqualsVariants = (X Is Y)
                End If
            End If
        Case vbNull
            EqualsVariants = IsNull(Y)
        Case vbEmpty
            EqualsVariants = IsEmpty(Y)
        Case VarType(Y)
            EqualsVariants = (X = Y)
    End Select
End Function
