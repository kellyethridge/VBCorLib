Attribute VB_Name = "Conversion"
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
' Module: Conversion
'
Option Explicit

Public Function CLngOrDefault(ByRef Value As Variant, ByVal Default As Long) As Long
    If IsMissing(Value) Then
        CLngOrDefault = Default
    Else
        CLngOrDefault = Value
    End If
End Function

Public Function CVarOrDefault(ByRef Value As Variant, ByRef Default As Variant) As Variant
    If IsMissing(Value) Then
        CVarOrDefault = Default
    Else
        CVarOrDefault = Value
    End If
End Function

Public Function CInt64(ByRef Value As Variant) As Int64
    Select Case VarType(Value)
        Case vbCurrency
            AsCurr(CInt64) = 0.0001@ * CCur(Int(Value))
        Case vbLong, vbInteger, vbByte
            CInt64.LowPart = CLng(Value)
            
            If CInt64.LowPart < 0 Then
                CInt64.HighPart = &HFFFFFFFF
            End If
        Case vbString
            CInt64 = Statics.Int64.Parse(CStr(Value))
        Case vbUserDefinedType
            If Not IsInt64(Value) Then _
                Throw New InvalidCastException
            
            CInt64 = Value
        Case Else
            Throw New InvalidCastException
    End Select
End Function

Public Function CUnk(ByVal Obj As IUnknown) As IUnknown
    Set CUnk = Obj
End Function


' This is a set of casting routines used by function delegation calls.
' They allow a virtual casting routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
'
Public Sub WidenLongToDouble(ByRef x As Double, ByRef y As Long)
    x = y
End Sub

Public Sub WidenLongToSingle(ByRef x As Single, ByRef y As Long)
    x = y
End Sub

Public Sub WidenLongToString(ByRef x As String, ByRef y As Long)
    x = y
End Sub

Public Sub WidenLongToCurrency(ByRef x As Currency, ByRef y As Long)
    x = y
End Sub

Public Sub WidenLongToVariant(ByRef x As Variant, ByRef y As Long)
    x = y
End Sub

Public Sub WidenIntegerToLong(ByRef x As Long, ByRef y As Integer)
    x = y
End Sub

Public Sub WidenIntegerToSingle(ByRef x As Single, ByRef y As Integer)
    x = y
End Sub

Public Sub WidenIntegerToDouble(ByRef x As Double, ByRef y As Integer)
    x = y
End Sub

Public Sub WidenIntegerToString(ByRef x As String, ByRef y As Integer)
    x = y
End Sub

Public Sub WidenIntegerToCurrency(ByRef x As Currency, ByRef y As Integer)
    x = y
End Sub

Public Sub WidenIntegerToVariant(ByRef x As Variant, ByRef y As Integer)
    x = y
End Sub

Public Sub WidenByteToInteger(ByRef x As Integer, ByRef y As Byte)
    x = y
End Sub

Public Sub WidenByteToLong(ByRef x As Long, ByRef y As Byte)
    x = y
End Sub

Public Sub WidenByteToSingle(ByRef x As Single, ByRef y As Byte)
    x = y
End Sub

Public Sub WidenByteToDouble(ByRef x As Double, ByRef y As Byte)
    x = y
End Sub

Public Sub WidenByteToString(ByRef x As String, ByRef y As Byte)
    x = y
End Sub

Public Sub WidenByteToCurrency(ByRef x As Currency, ByRef y As Byte)
    x = y
End Sub

Public Sub WidenByteToVariant(ByRef x As Variant, ByRef y As Byte)
    x = y
End Sub

Public Sub WidenSingleToDouble(ByRef x As Double, ByRef y As Single)
    x = y
End Sub

Public Sub WidenSingleToString(ByRef x As String, ByRef y As Single)
    x = y
End Sub

Public Sub WidenSingleToVariant(ByRef x As Variant, ByRef y As Single)
    x = y
End Sub

Public Sub WidenDateToDouble(ByRef x As Double, ByRef y As Date)
    x = y
End Sub

Public Sub WidenDateToString(ByRef x As String, ByRef y As Date)
    x = y
End Sub

Public Sub WidenDateToVariant(ByRef x As Variant, ByRef y As Date)
    x = y
End Sub

Public Sub WidenObjectToVariant(ByRef x As Variant, ByRef y As Object)
    Set x = y
End Sub

Public Sub WidenCurrencyToString(ByRef x As String, ByRef y As Currency)
    x = y
End Sub

Public Sub WidenCurrencyToVariant(ByRef x As Variant, ByRef y As Currency)
    x = y
End Sub

Public Sub WidenCurrencyToDouble(ByRef x As Double, ByRef y As Currency)
    x = y
End Sub

Public Sub WidenStringToVariant(ByRef x As Variant, ByRef y As String)
    x = y
End Sub

Public Sub WidenDoubleToString(ByRef x As String, ByRef y As Double)
    x = y
End Sub

Public Sub WidenDoubleToVariant(ByRef x As Variant, ByRef y As Double)
    x = y
End Sub

' Functions used to assign variants to narrower variables.
Public Sub NarrowVariantToLong(ByRef x As Long, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToInteger(ByRef x As Integer, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToDouble(ByRef x As Double, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToString(ByRef x As String, ByRef y As Variant)
    x = y
End Sub
Public Sub NarrowVariantToSingle(ByRef x As Single, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToByte(ByRef x As Byte, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToDate(ByRef x As Date, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToBoolean(ByRef x As Boolean, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToCurrency(ByRef x As Currency, ByRef y As Variant)
    x = y
End Sub

Public Sub NarrowVariantToObject(ByRef x As Object, ByRef y As Variant)
    Set x = y
End Sub
