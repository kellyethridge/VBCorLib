Attribute VB_Name = "modFormattable"
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
'
' Module: modFormattable
'

''
' Contains methods to support the formatted text output of values.
'
Option Explicit

''
' Converts a datatype value to a string representation using any
' supplied formatting or provider arguments.
'
' @param Value The value to convert to a string.
' @param Format Formatting information for converting the value.
' @param Provider A formatting provider to help custom formatting.
' @return A string representation of the value.
'
Public Function ToString(ByRef Value As Variant, Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim ValueType As Long
    ValueType = VarType(Value)
    
    If ValueType = vbVariant Then
        ValueType = MemLong(MemLong(VarPtr(Value) + VARIANTDATA_OFFSET)) And &HFF
    End If
    
    Select Case ValueType
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            Dim NumberFormatter As NumberFormatInfo
            If Not Provider Is Nothing Then
                Set NumberFormatter = Provider.GetFormat("numberformatinfo")
            End If
            
            If NumberFormatter Is Nothing Then
                Set NumberFormatter = NumberFormatInfo.CurrentInfo
            End If
            
            ToString = NumberFormatter.Format(Value, Format)
            
        Case vbDate
            Dim DateFormatter As DateTimeFormatInfo
            If Not Provider Is Nothing Then
                Set DateFormatter = Provider.GetFormat("datetimeformatinfo")
            End If
            
            If DateFormatter Is Nothing Then
                Set DateFormatter = DateTimeFormatInfo.CurrentInfo
            End If
            
            ToString = DateFormatter.Format(Value, Format)
            
        Case vbObject
            If Value Is Nothing Then
                ToString = ""
            ElseIf TypeOf Value Is IFormattable Then
                Dim Formattable As IFormattable
                Set Formattable = Value
                ToString = Formattable.ToString(Format, Provider)
            ElseIf TypeOf Value Is IObject Then
                Dim Obj As IObject
                Set Obj = Value
                ToString = Obj.ToString
            Else
                ToString = TypeName(Value)
            End If
        Case vbEmpty
            ToString = "Empty"
        Case vbNull
            ToString = "Null"
        Case vbMissing
            Exit Function
        Case Else
            ToString = Value
    End Select
End Function

