VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: BitConverter
'

''
' Provides a means to convert value datatypes to and from arrays of bytes.
'
' @remarks The <b>BitConverter</b> class cannot be instantiated. The methods
' are available through the class name itself.
' <pre>
' Dim b() As Byte
' b = BitConverter.GetBytes(&H12345678)
' </pre>
'
Option Explicit
Private Const ALPHA_OFFSET  As Long = 55

Private mBuffer As WordBuffer


''
' Converts the bit representation of a Double to a Currency.
'
' @param Value The Double bits to convert to Currency bits.
' @return The bit representation from the Double.
' @remarks This does not convert a Double value to that of a
' Currency value (like using CCur). This will copy the 64 bits
' of the Double variable to the 64 bits of the Currency variable.
'
Public Function DoubleToInt64Bits(ByVal Value As Double) As Currency
    DoubleToInt64Bits = AsCurr(Value)
End Function

''
' Converts the bit representation of a Currencty to a Double.
'
' @param Value The Currency bits to conver to Double bits.
' @return The bit representation from a Currency datatype.
' @remarks This does not convert a Currency to that of a
' Double value (like using CDbl). This will copy the 64 bits
' of the Currency variable to the 64 bits of the Double variable.
'
Public Function Int64BitsToDouble(ByVal Value As Currency) As Double
    Int64BitsToDouble = AsDouble(Value)
End Function

''
' Returns a byte array representation of the datatype value.
'
' @param value The datatype to be converted to a byte array.
' @return An array of bytes converted from the bits of the original datatype.
' @remarks <p>A type boolean is converted to one byte. If the boolean
' value is True, then the byte is 1, otherwise, it is 0.</p>
' <p>The byte array contains as many bytes as the length of the
' datatype being converted. A Long returns a 4-byte array, where as
' a Double will return an 8-byte array.</p>
' <p>The Decimal datatype is converted to 16 bytes with the same
' layout as used in .NET. The layout is different than that of VB.<br><br>
' VB Decimal layout<br>
' bytes 0-1: Variant datatype information<br>
' bytes 2: precision<br>
' bytes 3: sign (&h80 is negative)<br>
' bytes 4-7: the 32 highest bits in the 96bit value<br>
' bytes 8-11: the 32 lowest bits in the 96bit value<br>
' bytes 12-15: the 32 middle bits in the 96bit value<br>
' <br>
' .NET Decimal layout<br>
' bytes 0-3: the 32 lowest bits in the 96bit value<br>
' bytes 4-7: the 32 middle bits in the 96bit value<br>
' bytes 8-11: the 32 highest bits in the 96bit value<br>
' bytes 12-13: unused (zero)<br>
' bytes 14: precision<br>
' bytes 15: sign (&h80 is negative)<br>
'
Public Function GetBytes(ByRef Value As Variant) As Byte()
    Dim Ret() As Byte

    ' We use SafeArrayCreateVectorEx because the array descriptor
    ' and data are allocated all in one shot, where as using
    ' ReDim will cause two heap memory allocations.
    Select Case VarType(Value)
        Case vbLong
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 4)
            AsLong(Ret(0)) = CLng(Value)
            
        Case vbBoolean
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 1)
            If Value Then
                Ret(0) = 1
            End If
            
        Case vbInteger
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 2)
            AsWord(Ret(0)) = CInt(Value)
            
        Case vbByte
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 1)
            Ret(0) = CByte(Value)
            
        Case vbDouble, vbDate
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 8)
            AsDouble(Ret(0)) = CDbl(Value)
            
        Case vbCurrency
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 8)
            AsCurr(Ret(0)) = CCur(Value)
            
        Case vbSingle
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 4)
            AsSingle(Ret(0)) = CSng(Value)
            
        Case vbDecimal
            Dim ptrDec As Long
            Dim ptrRet As Long
            
            SAPtr(Ret) = SafeArrayCreateVectorEx(vbByte, 0, 16)
            
            If VariantType(Value) And VT_BYREF Then
                ptrDec = MemLong(VarPtr(Value) + 8)
            Else
                ptrDec = VarPtr(Value)
            End If
            ptrRet = VarPtr(Ret(0))
            
            ' VB Decimal layout
            ' bytes 0-1: Variant datatype information
            ' bytes 2-3: precision and sign bytes
            ' bytes 4-7: the 32 highest bits in the 96bit value
            ' bytes 8-11: the 32 lowest bits in the 96bit value
            ' bytes 12-15: the 32 middle bits in the 96bit value
            '
            ' .NET Decimal layout
            ' bytes 0-3: the 32 lowest bits in the 96bit value
            ' bytes 4-7: the 32 middle bits in the 96bit value
            ' bytes 8-11: the 32 highest bits in the 96bit value
            ' bytes 12-13: unused (zero)
            ' bytes 14-15: precision and sign bytes
            MemCurr(ptrRet) = MemCurr(ptrDec + 8)       ' map VB lowest and middle bits to .NET lowest and middle bits
            MemLong(ptrRet + 8) = MemLong(ptrDec + 4)   ' map VB highest bits to .NET highest bits
            ' In .NET the precision and sign bytes are last.
            ' we don't copy the type information from the variant.
            MemWord(ptrRet + 14) = MemWord(ptrDec + 2)
        Case Else
            Throw Cor.NewNotSupportedException("Only intrinsic value types are supported.")
    End Select
    GetBytes = Ret
End Function

''
' Converts an array of bytes to a Boolean value.
'
' @param Value The array of bytes to create a boolean from.
' @param StartIndex The byte in the array to create the boolean from.
' @return Boolean representation of the byte used in the array.
' @remarks The boolean value is created from one byte. If the byte is
' a non-zero value, True is returned, otherwise, False is returned.
'
Public Function ToBoolean(ByRef Value() As Byte, ByVal StartIndex As Long) As Boolean
    ValidateArrayArguments Value, StartIndex, 1
    ToBoolean = (Value(StartIndex) > 0)
End Function

''
' Converts an array of bytes to a Double value.
'
' @param Value The array of bytes to create a double from.
' @param StartIndex The starting byte in the array to create a double from.
' @return Double datatype representation of the 8 bytes used in the array.
'
Public Function ToDouble(ByRef Value() As Byte, ByVal StartIndex As Long) As Double
    ValidateArrayArguments Value, StartIndex, 8
    ToDouble = AsDouble(Value(StartIndex))
End Function

''
' Converts an array of bytes to a Date value.
'
' @param Value The array of bytes to create a date from.
' @param StartIndex The starting byte in the array to create a date from.
' @return Date datatype representation of the 8 bytes used in the array.
' @remarks Even though a Date is represented using a Double, it may still
' raise an error due to the 8 bytes used represent an invalid date range.
'
Public Function ToDate(ByRef Value() As Byte, ByVal StartIndex As Long) As Date
    ValidateArrayArguments Value, StartIndex, 8
    ToDate = AsDouble(Value(StartIndex))
End Function

''
' Converts an array of bytes to an Integer value.
'
' @param Value The array of bytes used to create an integer from.
' @param StartIndex The starting byte in the array to create an integer from.
' @return Integer datatype representation of the 2 bytes used in the array.
'
Public Function ToInteger(ByRef Value() As Byte, ByVal StartIndex As Long) As Integer
    ValidateArrayArguments Value, StartIndex, 2
    ToInteger = AsWord(Value(StartIndex))
End Function

''
' Converts an array of bytes to a Long value.
'
' @param Value The array of bytes used to create a Long from.
' @param StartIndex The starting byte in the array to create a Long from.
' @return Long datatype representation of the 4 bytes used in the array.
'
Public Function ToLong(ByRef Value() As Byte, ByVal StartIndex As Long) As Long
    ValidateArrayArguments Value, StartIndex, 4
    ToLong = AsLong(Value(StartIndex))
End Function

''
' Converts an array of bytes to a Single value.
'
' @param Value The array of bytes used to create a Single from.
' @param StartIndex The starting byte in the array to create a Single from.
' @return Single datatype representation of the 4 bytes used in the array.
'
Public Function ToSingle(ByRef Value() As Byte, ByVal StartIndex As Long) As Single
    ValidateArrayArguments Value, StartIndex, 4
    ToSingle = AsSingle(Value(StartIndex))
End Function

''
' Converts an array of bytes to a Currency value.
'
' @param Value The array of bytes used to create a Currency from.
' @param StartIndex The starting byte in the array to create a Currency from.
' @return Currency datatype representation of the 8 bytes used in the array.
'
Public Function ToCurrency(ByRef Value() As Byte, ByVal StartIndex As Long) As Currency
    ValidateArrayArguments Value, StartIndex, 8
    ToCurrency = AsCurr(Value(StartIndex))
End Function

''
' Converts an array of bytes do a Variant Decimal value.
'
' @param Value The array of bytes used to create a Decimal value. Must be atleast 16 bytes in length.
' @param StartIndex The starting index within the array to be converting to a Decimal value.
' @return The converted Decimal value.
' @remarks <p>The Decimal datatype is created from 16 bytes with the same
' layout as used in .NET. The layout is different than that of VB.<br><br>
' VB Decimal layout<br>
' bytes 0-1: Variant datatype information<br>
' bytes 2: precision<br>
' bytes 3: sign (&h80 is negative)<br>
' bytes 4-7: the 32 highest bits in the 96bit value<br>
' bytes 8-11: the 32 lowest bits in the 96bit value<br>
' bytes 12-15: the 32 middle bits in the 96bit value<br>
' <br>
' .NET Decimal layout<br>
' bytes 0-3: the 32 lowest bits in the 96bit value<br>
' bytes 4-7: the 32 middle bits in the 96bit value<br>
' bytes 8-11: the 32 highest bits in the 96bit value<br>
' bytes 12-13: unused (zero)<br>
' bytes 14: precision<br>
' bytes 15: sign (&h80 is negative)<br><br>
' If the sign byte is non-zero, then the Decimal will be set as negative.<br>
'
Public Function ToDecimal(ByRef Value() As Byte, ByVal StartIndex As Long) As Variant
    ValidateArrayArguments Value, StartIndex, 16
    VariantType(ToDecimal) = vbDecimal
    
    ' check precision
    If Value(StartIndex + 14) > 28 Then _
        Throw Cor.NewOverflowException("Decimal precision must be from 0 to 28 places.")
    
    Dim ptr As Long
    ptr = VarPtr(ToDecimal)
    
    If Value(StartIndex + 15) <> 0 Then MemByte(ptr + 3) = &H80  ' set sign
    
    MemByte(ptr + 2) = Value(StartIndex + 14)            ' set precision
    MemLong(ptr + 4) = AsLong(Value(StartIndex + 8))     ' set highest 32 bits
    MemCurr(ptr + 8) = AsCurr(Value(StartIndex + 0))     ' set lowest and middle 32bits
End Function

''
' Converts an array of bytes to a string of hexidecimal notations.
'
' @param Value The array of bytes used to create a string of hexidecimal values.
' @param StartIndex The starting byte to begin creating hexidecimal values from in the array.
' @param Count The number of bytes to be converted to a hexidecimal notation.
' @return String containing hexidecimal notations for each byte, separated by hyphens.
' @remarks Each byte to be converted to string is converted into a hexidecimal representation.
' For example a byte value of 160 would become 'A0' in the return string. Each byte is separated
' by a hyphen when more than one byte is being converted. An example is a Long of &hABCDEF01 would
' become '01-EF-CD-AB'.
'
Public Function ToString(ByRef Value() As Byte, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As String
    If cArray.IsNull(Value) Then _
        Throw Cor.NewArgumentNullException("Value", Environment.GetResourceString(ArgumentNull_Array))
    
    Dim ElemIndex As Long
    ElemIndex = GetOptionalLong(StartIndex, LBound(Value))
    
    Dim ElemCount As Long
    ElemCount = GetOptionalLong(Count, UBound(Value) - ElemIndex + 1)
    
    ValidateArrayArguments Value, ElemIndex, ElemCount
    
    Dim Ret As String
    Ret = SysAllocStringLen(0, ElemCount * 3 - 1)
    mBuffer.SA.pvData = StrPtr(Ret)
    
    Dim i       As Long
    Dim b       As Long
    Dim Pos     As Long
    Dim nibble  As Long
    Dim Offset  As Long
    
    Offset = 1
    Do While i < ElemCount
        If i > 0 Then mBuffer.Data(Pos) = vbMinus: Pos = Pos + 1
        b = Value(ElemIndex + i)
        Do
            nibble = b And &HF
            If nibble > 9 Then
                mBuffer.Data(Pos + Offset) = nibble + ALPHA_OFFSET
            Else
                mBuffer.Data(Pos + Offset) = nibble + vbZero
            End If
            b = (b And &HF0&) \ &H10&
            Offset = 1 - Offset
        Loop While Offset = 0
        Pos = Pos + 2
        i = i + 1
    Loop
    
    ToString = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateArrayArguments(ByRef Value() As Byte, ByVal StartIndex As Long, ByVal Count As Long)
    If IsNullByteArray(Value) Then _
        Throw Cor.NewArgumentNullException("Value")
    If StartIndex < LBound(Value) Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", Message:=Resources.GetString(ArgumentOutOfRange_ArrayLB))
    If StartIndex > UBound(Value) - Count + 1 Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", Message:=Resources.GetString(ArgumentOutOfRange_Index))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitWordBuffer mBuffer, 0, &H7FFFFFFF
End Sub

