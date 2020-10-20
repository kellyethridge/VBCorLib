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
' Converts base data types to an array of bytes, and an array of bytes to base data types.
'
' @include "..\DocIncludes\BitConverter.html"
'@Folder("CorLib.System")
Option Explicit


''
' Converts the specified double-precision floating point number to 64-bit signed integer.
'
' @param Value The number to convert.
' @return A 64-bit signed integer whose value is equivalent to <i>Value</i>.
' @remarks This copies the 64-bit representation of a double-precision floating point number
' to a 64-bit integer. It does not attempt to convert the value similar to using <a href="PublicFunctions.CInt64.html">CInt64</a>.
Public Function DoubleToInt64Bits(ByVal Value As Double) As Int64
    AsDouble(DoubleToInt64Bits) = Value
End Function

''
' Converts the specified 64-bit signed integer to a double-precision floating point number.
'
' @param Value The number to convert.
' @return A double-precision floating point number whose value is equivalent to <i>Value</i>.
' @remarks This copies the 64-bits of the <i>Value</i> parameter to double-precision value.
' <p>Typically, <i>Value</i> is an integer that is returned by the <a href="BitConverter.DoubleToInt64Bits.html">DoubleToInt64Bits</a> method.</p>
Public Function Int64BitsToDouble(ByRef Value As Int64) As Double
    Int64BitsToDouble = AsDouble(Value)
End Function

''
' Converts the specified single-precision floating point number to 32-bit signed integer.
'
' @param Value The number to convert.
' @return A 32-bit signed integer whose value is equivalent to <i>Value</i>.
' @remarks This copies the 32-bit representation of a single-precision floating point number
' to a 32-bit integer. It does not attempt to convert the value similar to using <b>CLng</b>.
Public Function SingleToInt32Bits(ByVal Value As Single) As Long
    SingleToInt32Bits = AsLong(Value)
End Function

''
' Converts the specified 32-bit signed integer to a single-precision floating point number.
'
' @param Value The number to convert.
' @return A single-precision floating point number whose value is equivalent to <i>Value</i>.
' @remarks This copies the 32-bits of the <i>Value</i> parameter to single-precision value.
' <p>Typically, <i>Value</i> is an integer that is returned by the <a href="BitConverter.SingleToInt32Bits.html">SingleToInt32Bits</a> method.</p>
Public Function Int32BitsToSingle(ByVal Value As Long) As Single
    Int32BitsToSingle = AsSingle(Value)
End Function

''
' Converts the specified currency number to 64-bit signed integer.
'
' @param Value The number to convert.
' @return A 64-bit signed integer whose value is equivalent to <i>Value</i>.
' @remarks This copies the 64-bit representation of a currency number to a 64-bit integer.
' It does not attempt to convert the value similar to using <a href="PublicFunctions.CInt64.html">CInt64</a>.
Public Function CurrencyToInt64Bits(ByVal Value As Currency) As Int64
    AsCurr(CurrencyToInt64Bits) = Value
End Function

''
' Converts the specified 64-bit signed integer to a currency number.
'
' @param Value The number to convert.
' @return A currency number whose value is equivalent to <i>Value</i>.
' @remarks This copies the 64-bits of the <i>Value</i> parameter to currency value.
' <p>Typically, <i>Value</i> is an integer that is returned by the <a href="BitConverter.CurrencyToInt64Bits.html">CurrencyToInt64Bits</a> method.</p>
Public Function Int64BitsToCurrency(ByRef Value As Int64) As Currency
    Int64BitsToCurrency = AsCurr(Value)
End Function

''
' Returns a byte array representation of the datatype value.
'
' @param value The datatype to be converted to a byte array.
' @return An array of bytes converted from the bits of the original datatype.
' @remarks <p>A type boolean is converted to one byte. If the boolean value is True, then the byte is 1, otherwise, it is 0.</p>
' <p>The byte array contains as many bytes as the length of the datatype being converted. A Long returns a 4-byte array, where as a Double will return an 8-byte array.</p>
' <p>The following is a list of types that are supported.</p>
' <ul>
'   <li>Boolean</li>
'   <li>Byte</li>
'   <li>Integer</li>
'   <li>Long</li>
'   <li><a href="Int64Static.Int64.html">Int64</a></li>
'   <li>Single</li>
'   <li>Double</li>
'   <li>Currency</li>
'   <li>Date</li>
'   <li>Decimal</li>
' </ul>
' <p>The <b>Decimal</b> type is converted to 16 bytes with the same layout as used in .NET. The layout is different than that of Visual Basic.</p>
' <p>The differing layouts is shown in the following table.</p>
' <table style="width: auto">
'   <tr>
'       <th>Framework</th><th>Layout</th>
'   </tr>
'   <tr>
'        <td>Visual Basic</td>
'        <td>
'           bytes 0-1: Variant type information (Decimal has value 14)<br/>
'           bytes 2: precision<br/>
'           bytes 3: sign (&h80 is negative)<br/>
'           bytes 4-7: the 32 highest bits in the 96bit value<br/>
'           bytes 8-11: the 32 lowest bits in the 96bit value<br/>
'           bytes 12-15: the 32 middle bits in the 96bit value
'        </td>
'   </tr>
'    <tr>
'        <td>.NET</td>
'        <td>
'           bytes 0-3: the 32 lowest bits in the 96bit value<br/>
'           bytes 4-7: the 32 middle bits in the 96bit value<br/>
'           bytes 8-11: the 32 highest bits in the 96bit value<br/>
'           bytes 12-13: unused (zero)<br/>
'           bytes 14: precision<br/>
'           bytes 15: sign (&h80 is negative)
'        </td>
'    </tr>
' </table>
Public Function GetBytes(ByRef Value As Variant) As Byte()
    Dim Result() As Byte

    ' We use SafeArrayCreateVectorEx because the array descriptor
    ' and data are allocated all in one shot, where as using
    ' ReDim will cause two heap memory allocations.
    Select Case VarType(Value)
        Case vbLong
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfLong)
            AsLong(Result(0)) = CLng(Value)
        Case vbBoolean
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfByte)
            
            If Value Then
                Result(0) = 1
            End If
        Case vbInteger
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfInteger)
            AsWord(Result(0)) = CInt(Value)
        Case vbByte
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfByte)
            Result(0) = CByte(Value)
        Case vbDouble
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfDouble)
            AsDouble(Result(0)) = CDbl(Value)
        Case vbDate
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfCurrency)
            AssignQWord CorDateTime.GetTotalMilliseconds(Value), Result(0)
        Case vbCurrency
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfCurrency)
            AsCurr(Result(0)) = CCur(Value)
        Case vbSingle
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfSingle)
            AsSingle(Result(0)) = CSng(Value)
        Case vbUserDefinedType
            If Not IsInt64(Value) Then _
                Throw Cor.NewNotSupportedException("Only intrinsic value types are supported.")
                
            Dim Ptr As Long
            Ptr = DataPtr(Value)
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfInt64)
            AssignQWord ByVal Ptr, Result(0)
        Case vbDecimal
            Dim PtrSrc As Long
            Dim PtrDst As Long
            
            SAPtr(Result) = SafeArrayCreateVectorEx(vbByte, 0, vbSizeOfDecimal)
            PtrDst = VarPtr(Result(0))
            PtrSrc = VarPtr(Value)
            
            ' this would be an extremely rare occurrence since a decimal
            ' datatype can only exist within a variant variable. The passed
            ' in parameter would literally need to be a second variant
            ' variable containing a reference to a decimal variant variable.
            If VariantType(Value) And VT_BYREF Then
                ' if a value was passed by reference we need to
                ' dereference the memory address stored in the
                ' passed in variant value.
                PtrSrc = MemLong(PtrSrc + VARIANTDATA_OFFSET)
            End If
            
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
            AssignQWord ByVal PtrSrc + 8, ByVal PtrDst      ' map VB lowest and middle bits to .NET lowest and middle bits
            AssignDWord ByVal PtrSrc + 4, ByVal PtrDst + 8  ' map VB highest bits to .NET highest bits
            AssignWord ByVal PtrSrc + 2, ByVal PtrDst + 14  ' In .NET the precision and sign bytes are last. We don't copy the type information from the variant.
        Case Else
            Throw Cor.NewNotSupportedException("Only intrinsic value types are supported.")
    End Select
    
    GetBytes = Result
End Function

''
' Converts an array of bytes to a Boolean value.
'
' @param Value The array of bytes to create a boolean from.
' @param StartIndex The byte in the array to create the boolean from.
' @return Boolean representation of the byte used in the array.
' @remarks The boolean value is created from one byte. If the byte is
' a non-zero value, <b>True</b> is returned, otherwise, <b>False</b> is returned.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
' </table>
Public Function ToBoolean(ByRef Value() As Byte, ByVal StartIndex As Long) As Boolean
    ValidateArrayArguments Value, StartIndex, vbSizeOfByte
    ToBoolean = (Value(StartIndex) > 0)
End Function

''
' Converts an array of bytes to a Double value.
'
' @param Value The array of bytes to create a double from.
' @param StartIndex The starting byte in the array to create a double from.
' @return Double data-type representation of the 8 bytes used in the array.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 7,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
' </table>
Public Function ToDouble(ByRef Value() As Byte, ByVal StartIndex As Long) As Double
    ValidateArrayArguments Value, StartIndex, vbSizeOfDouble
    ToDouble = AsDouble(Value(StartIndex))
End Function

''
' Converts an array of bytes to a Date value.
'
' @param Value The array of bytes to create a date from.
' @param StartIndex The starting byte in the array to create a date from.
' @return Date datatype representation of the 8 bytes used in the array.
' @remarks The date type is stored as ticks in a 64-bit signed integer. The Visual Basic Date type cannot represent the entire range storable in 64-bits.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 7,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
'   <tr>
'        <td><a href="OverflowException.html">OverflowException</a></td><td>The 64-bit value is outside the valid range of the Visual Basic Date type.</td>
'   </tr>
' </table>
Public Function ToDate(ByRef Value() As Byte, ByVal StartIndex As Long) As Date
    Dim Milliseconds As Currency
    
    Milliseconds = ToCurrency(Value, StartIndex)
    ToDate = CorDateTime.GetOADate(Milliseconds)
End Function

''
' Converts an array of bytes to an Integer value.
'
' @param Value The array of bytes used to create an integer from.
' @param StartIndex The starting byte in the array to create an integer from.
' @return Integer data-type representation of the 2 bytes used in the array.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 1,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
' </table>
Public Function ToInt16(ByRef Value() As Byte, ByVal StartIndex As Long) As Integer
    ValidateArrayArguments Value, StartIndex, vbSizeOfInteger
    ToInt16 = AsWord(Value(StartIndex))
End Function

''
' Converts an array of bytes to a Long value.
'
' @param Value The array of bytes used to create a Long from.
' @param StartIndex The starting byte in the array to create a Long from.
' @return Long data-type representation of the 4 bytes used in the array.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 3,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
' </table>
Public Function ToInt32(ByRef Value() As Byte, ByVal StartIndex As Long) As Long
    ValidateArrayArguments Value, StartIndex, vbSizeOfLong
    ToInt32 = AsLong(Value(StartIndex))
End Function

''
' Returns a 64-bit signed integer converted from eight bytes at a specified position in a byte array.
'
' @param Value An array of bytes.
' @param StartIndex The starting position within <i>Value</i>.
' @return A 64-bit signed integer formed by eight bytes beginning at <i>StartIndex</i>.
' @remarks The <b>ToInt64</b> method converts the bytes from index <i>StartIndex</i> to <i>StartIndex</i> + 7 to a <a href="Int64Static.Int64.html">Int64</a> value. The order of bytes in the array must be little-endian.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 7,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
' </table>
Public Function ToInt64(ByRef Value() As Byte, ByVal StartIndex As Long) As Int64
    ValidateArrayArguments Value, StartIndex, vbSizeOfInt64
    AssignQWord Value(StartIndex), ToInt64
End Function

''
' Converts an array of bytes to a Single value.
'
' @param Value The array of bytes used to create a Single from.
' @param StartIndex The starting byte in the array to create a Single from.
' @return Single data-type representation of the 4 bytes used in the array.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 3,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
' </table>
Public Function ToSingle(ByRef Value() As Byte, ByVal StartIndex As Long) As Single
    ValidateArrayArguments Value, StartIndex, vbSizeOfSingle
    ToSingle = AsSingle(Value(StartIndex))
End Function

''
' Converts an array of bytes to a Currency value.
'
' @param Value The array of bytes used to create a Currency from.
' @param StartIndex The starting byte in the array to create a Currency from.
' @return Currency data-type representation of the 8 bytes used in the array.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 7,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.</td>
'    </tr>
' </table>
Public Function ToCurrency(ByRef Value() As Byte, ByVal StartIndex As Long) As Currency
    ValidateArrayArguments Value, StartIndex, vbSizeOfCurrency
    ToCurrency = AsCurr(Value(StartIndex))
End Function

''
' Converts an array of bytes do a Variant Decimal value.
'
' @param Value The array of bytes used to create a Decimal value. Must be atleast 16 bytes in length.
' @param StartIndex The starting index within the array to be converting to a Decimal value.
' @return The converted Decimal value.
' @remarks The Decimal type is created from 16 bytes with the same layout as used in .NET. The layout is different than that of Visual Basic.
' <p>The differing layouts is shown in the following table.</p>
' <table style="width: auto">
'   <tr>
'       <th>Framework</th><th>Layout</th>
'   </tr>
'   <tr>
'        <td>Visual Basic</td>
'        <td>
'           bytes 0-1: Variant type information (Decimal has value 14)<br/>
'           bytes 2: precision<br/>
'           bytes 3: sign (&h80 is negative)<br/>
'           bytes 4-7: the 32 highest bits in the 96bit value<br/>
'           bytes 8-11: the 32 lowest bits in the 96bit value<br/>
'           bytes 12-15: the 32 middle bits in the 96bit value
'        </td>
'   </tr>
'    <tr>
'        <td>.NET</td>
'        <td>
'           bytes 0-3: the 32 lowest bits in the 96bit value<br/>
'           bytes 4-7: the 32 middle bits in the 96bit value<br/>
'           bytes 8-11: the 32 highest bits in the 96bit value<br/>
'           bytes 12-13: unused (zero)<br/>
'           bytes 14: precision<br/>
'           bytes 15: sign (&h80 is negative)
'        </td>
'    </tr>
' </table>
' If the sign byte is non-zero, then the Decimal will be set as negative.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>StartIndex</i> is greater than or equal to the length of <i>Value</i> minus 15,
'            and is less than or equal to the length of <i>Value</i> minus 1.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is uninitialized.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td><i>StartIndex</i> is less than zero or greater than the length of <i>Value</i> minus 1.<br/>
'            -or-<br/>
'            The precision is greater than 28.
'        </td>
'    </tr>
' </table>
Public Function ToDecimal(ByRef Value() As Byte, ByVal StartIndex As Long) As Variant
    ValidateArrayArguments Value, StartIndex, vbSizeOfDecimal
    VariantType(ToDecimal) = vbDecimal
    
    ' check precision
    If Value(StartIndex + 14) > 28 Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_DecimalScale
    
    Dim Ptr As Long
    Ptr = VarPtr(ToDecimal)
    
    If Value(StartIndex + 15) <> 0 Then MemByte(Ptr + 3) = &H80  ' set sign
    
    MemByte(Ptr + 2) = Value(StartIndex + 14)            ' set precision
    MemLong(Ptr + 4) = AsLong(Value(StartIndex + 8))     ' set highest 32 bits
    MemCurr(Ptr + 8) = AsCurr(Value(StartIndex + 0))     ' set lowest and middle 32bits
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
    Const ALPHA_OFFSET As Long = 55

    If SAPtr(Value) = vbNullPtr Then _
        Error.ArgumentNull "Value", ArgumentNull_Array
    
    Dim Range As ListRange
    Range = MakeDefaultStepRange(StartIndex, LBound(Value), Count, Len1D(Value), Parameter_StartIndex)
    
    ValidateArrayArguments Value, Range.Index, Range.Count
    
    Dim Ret As String
    Ret = SysAllocStringLen(0, Range.Count * 3 - 1)
    
    Dim Chars() As Integer
    Chars = AllocChars(Ret)
    
    Dim i       As Long
    Dim b       As Long
    Dim pos     As Long
    Dim Nibble  As Long
    Dim Offset  As Long
    
    Offset = 1
    Do While i < Range.Count
        If i > 0 Then
            Chars(pos) = vbMinusChar
            pos = pos + 1
        End If
        
        b = Value(Range.Index + i)
        Do
            Nibble = b And &HF
            If Nibble > 9 Then
                Chars(pos + Offset) = Nibble + ALPHA_OFFSET
            Else
                Chars(pos + Offset) = Nibble + vbZeroChar
            End If
            b = (b And &HF0&) \ &H10&
            Offset = 1 - Offset
        Loop While Offset = 0
        
        pos = pos + 2
        i = i + 1
    Loop
    
    FreeChars Chars
    ToString = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateArrayArguments(ByRef Value() As Byte, ByVal StartIndex As Long, ByVal Count As Long)
    If SAPtr(Value) = vbNullPtr Then
        Error.ArgumentNull "Value", ArgumentNull_Array
    End If
    If StartIndex < LBound(Value) Then
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_ArrayLB
    End If
    If StartIndex > UBound(Value) - Count + 1 Then
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    End If
End Sub

