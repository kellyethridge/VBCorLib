VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Int64Static"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: Int64Static
'

''
' Provides methods for manipulating a 64-bit signed integer.
'
Option Explicit

''
' Represents a 64-bit signed integer.
'
' @param LowPart The least significant 32-bits of the integer value.
' @param HighPart The most significant 32-bits of the integer value.
Public Type Int64
    LowPart     As Long ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
    HighPart    As Long
End Type


''
' Represents the smallest possible value of an <a href="Int64Static.Int64.html">Int64</a>. This field is constant.
'
' @return The minimum value constant.
' @remarks The value of this constant is negative 9,223,372,036,854,775,808; that is, hexadecimal 0x8000000000000000.
Public Property Get MinValue() As Int64
    MinValue.HighPart = &H80000000
End Property

''
' Represents the largest possible value of an <a href="Int64Static.Int64.html">Int64</a>. This field is constant.
'
' @return The maximum value constant.
' @remarks The value of this constant is 9,223,372,036,854,775,807; that is, hexadecimal 0x7FFFFFFFFFFFFFFF.
Public Property Get MaxValue() As Int64
    MaxValue.HighPart = &H7FFFFFFF
    MaxValue.LowPart = &HFFFFFFFF
End Property

''
' Returns an <a href="Int64Static.Int64.html">Int64</a> value representing zero.
'
' @return The <a href="Int64Static.Int64.html">Int64</a> with zero value.
Public Property Get Zero() As Int64
    ' simply returns uninitialized value
End Property

''
' Returns an <a href="Int64Static.Int64.html">Int64</a> value representing one.
'
' @return The <a href="Int64Static.Int64.html">Int64</a> with a value of one.
Public Property Get One() As Int64
    One.LowPart = 1
End Property

''
' Returns an <a href="Int64Static.Int64.html">Int64</a> value represnting minus-one.
'
' @return The <a href="Int64Static.Int64.html">Int64</a> with a value of minus-one.
Public Property Get MinusOne() As Int64
    MinusOne.LowPart = &HFFFFFFFF
    MinusOne.HighPart = &HFFFFFFFF
End Property

''
' Returns if value is equal to zero.
'
' @param Value The <a href="Int64Static.Int64.html">Int64</a> value to check.
' @return <b>True</b> if the value is equal to zero; <b>False</b> otherwise.
Public Function IsZero(ByRef Value As Int64) As Boolean
    IsZero = (Value.LowPart Or Value.HighPart) = 0
End Function

''
' Returns if value is equal to one.
'
' @param Value The <a href="Int64Static.Int64.html">Int64</a> value to check.
' @return <b>True</b> if the value is equal to one; <b>False</b> otherwise.
Public Function IsOne(ByRef Value As Int64) As Boolean
    IsOne = (Value.LowPart = 1) And (Value.HighPart = 0)
End Function

''
' Returns if value is equal to minue one.
'
' @param Value The <a href="Int64Static.Int64.html">Int64</a> value to check.
' @return <b>True</b> if the value is equal to minus one; <b>False</b> otherwise.
Public Function IsMinusOne(ByRef Value As Int64) As Boolean
    IsMinusOne = (Value.LowPart = -1) And (Value.HighPart = -1)
End Function

''
' Returns an integer indicating the sign of the value.
'
' @param Value The number to return the sign of.
' @return The sign of the number.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Meaning</th>
'   </tr>
'    <tr>
'        <td>1</td><td><i>Value</i> is greater than zero.</td>
'    </tr>
'    <tr>
'        <td>0</td><td><i>Value</i> equals zero.</td>
'    </tr>
'   <tr>
'        <td>-1</td><td><i>Value</i> is less than zero.</td>
'   </tr>
' </table>
Public Function Sign(ByRef Value As Int64) As Long
    If Value.HighPart And &H80000000 Then
        Sign = -1
    ElseIf Not IsZero(Value) Then
        Sign = 1
    End If
End Function

''
' Determines of two <a href="Int64Static.Int64.html">Int64</a> values are equal.
'
' @param Int64A The first value to compare.
' @param Int64B The second value to compare.
' @return <b>True</b> if the values are equal; otherwise, <b>False</b>.
Public Function Equals(ByRef Int64A As Int64, ByRef Int64B As Int64) As Boolean
    Equals = (Int64A.HighPart = Int64B.HighPart) And (Int64A.LowPart = Int64B.LowPart)
End Function

''
' Returns the hash code for this instance.
'
' @param Value The 64-bit integer value to get a hash code for.
' @return A 32-bit signed integer hash code.
Public Function GetHashCode(ByRef Value As Int64) As Long
    GetHashCode = Value.LowPart Xor Value.HighPart
End Function

''
' Converts the numeric value of this instance to its equivalent string representation using the specified format and culture-specific format information.
'
' @param Value The 64-bit integer to be converted to a string.
' @param Format A numeric format string.
' @param Provider An object that supplies culture-specific formatting information about this instance.
' @return The string representation of the value of this instance as specified by <i>Format</i> and <i>Provider</i>.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="FormatException.html">FormatException</a></td><td><i>Format</i> is invalid or not supported.</td>
'   </tr>
' </table>
Public Function ToString(ByRef Value As Int64, Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim FormatInfo  As NumberFormatInfo
    
    Set FormatInfo = NumberFormatInfo.GetInstance(Provider)
    ToString = FormatInfo.FormatInt64(Value, Format)
End Function


''
' Compares two <a href="Int64Static.Int64.html">Int64</a> values and returns a value indicating whether one is less than, equal to, or greater than the other.
'
' @param a The first <a href="Int64Static.Int64.html">Int64</a> to compare.
' @param b The second <a href="Int64Static.Int64.html">Int64</a> to compare.
' @return A signed integer that indicates the relative values of x and y, as shown in the following table.
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Meaning</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>X</i> is less than <i>Y</i>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>X</i> equals <i>Y</i>.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>X</i> is greater than <i>Y</i>.</td>
'    </tr>
' </table>
Public Function Compare(ByRef a As Int64, ByRef b As Int64) As Long
    If a.HighPart < b.HighPart Then
        Compare = -1
    ElseIf a.HighPart > b.HighPart Then
        Compare = 1
    ElseIf a.LowPart < 0 Then
        If b.LowPart < 0 Then
            If a.LowPart < b.LowPart Then
                Compare = -1
            ElseIf a.LowPart > b.LowPart Then
                Compare = 1
            End If
        Else
            Compare = 1
        End If
    ElseIf b.LowPart < 0 Then
        Compare = -1
    ElseIf a.LowPart < b.LowPart Then
        Compare = -1
    ElseIf a.LowPart > b.LowPart Then
        Compare = 1
    End If
End Function

''
' Returns a new <a href="Int64Static.Int64.html">Int64</a> whose value is the negated value of <i>Value</i>.
'
' @param Value The value to be negated.
' @return The negated value of <i>Value</i>.
Public Function Negate(ByRef Value As Int64) As Int64
    Negate.HighPart = Not Value.HighPart
    Negate.LowPart = Not Value.LowPart
    Negate = Add(Negate, One)
End Function

''
' Returns a new <a href="Int64Static.Int64.html">Int64</a> whose value is the absolute value of <i>Value</i>.
'
' @param Value The value to return the absolute value of.
' @return The absolute value of <i>Value</i>.
Public Function Absolute(ByRef Value As Int64) As Int64
    If Value.HighPart And &H80000000 Then
        Absolute = Negate(Value)
    Else
        Absolute = Value
    End If
End Function

''
' Returns the sum of two <a href="Int64Static.Int64.html">Int64</a> values.
'
' @param a The first value to be added.
' @param b The second value to be added.
' @return The sum of <i>a</b> and <i>b</i>.
Public Function Add(ByRef a As Int64, ByRef b As Int64) As Int64
    AsCurr(Add) = AsCurr(a) + AsCurr(b)
End Function

''
' Returns the difference of two <a href="Int64Static.Int64.html">Int64</a> values.
'
' @param a The value to be subtracted from.
' @param b The value to be subtracted.
' @return The difference of <i>a</i> and <i>b</i>.
Public Function Subtract(ByRef a As Int64, ByRef b As Int64) As Int64
    AsCurr(Subtract) = AsCurr(a) - AsCurr(b)
End Function

''
' Multiplies two <a href="Int64Static.Int64.html">Int64</a> values returning the product.
'
' @param a The first value to be multiplied.
' @param b The second value to be multiplied.
' @return The product of <i>a</i> and <i>b</i>.
Public Function Multiply(ByRef a As Int64, ByRef b As Int64) As Int64
    If a.HighPart And &H80000000 Then
        If b.HighPart And &H80000000 Then
            MultiplyPositives Negate(a), Negate(b), Multiply
        Else
            MultiplyByNegative b, a, Multiply
        End If
    ElseIf b.HighPart And &H80000000 Then
        MultiplyByNegative a, b, Multiply
    Else
        MultiplyPositives a, b, Multiply
    End If
End Function

''
' Calculates the quotient of two 64-bit signed integers.
'
' @param Dividend The value to be divided.
' @param Divisor The value to divide by.
' @return The integer quotient of the division.
' @remarks This method performs integer division, discarding the remainder.
Public Function Divide(ByRef Dividend As Int64, ByRef Divisor As Int64) As Int64
    Dim Remainder As Int64
    
    Divide = DivRem(Dividend, Divisor, Remainder)
End Function

''
' Calculates the remainder of two 64-bit signed integers.
'
' @param Dividend The value to be divided.
' @param Divisor The value to divide by.
' @return Returns the remainder portion of the division operation.
Public Function Modulus(ByRef Dividend As Int64, ByRef Divisor As Int64) As Int64
    DivRem Dividend, Divisor, Modulus
End Function

''
' Calculates the quotient of two 64-bit signed integers and also returns the remainder in an output parameter.
'
' @param Dividend The value to be divided.
' @param Divisor The value to divide by.
' @param Remainder The remainder of the division.
' @return The integer quotient of the division.
Public Function DivRem(ByRef Dividend As Int64, ByRef Divisor As Int64, ByRef Remainder As Int64) As Int64
    If IsZero(Divisor) Then _
        Throw New DivideByZeroException
    
    Remainder.HighPart = 0
    Remainder.LowPart = 0
    
    ' perform a couple of quick checks that have an easy result.
    If IsZero(Dividend) Then
        Exit Function
    ElseIf IsOne(Divisor) Then
        DivRem = Dividend
        Exit Function
    End If
    
    If Divisor.HighPart And &H80000000 Then
        If Dividend.HighPart And &H80000000 Then
            DivideCore Negate(Dividend), Negate(Divisor), DivRem, Remainder
        Else
            DivideCore Dividend, Negate(Divisor), DivRem, Remainder
            DivRem = Negate(DivRem)
        End If
    ElseIf Dividend.HighPart And &H80000000 Then
        DivideCore Negate(Dividend), Divisor, DivRem, Remainder
        DivRem = Negate(DivRem)
    Else
        DivideCore Dividend, Divisor, DivRem, Remainder
    End If
End Function

''
' Performs a bitwise AND of two <a href="Int64Static.Int64.html">Int64</a> values.
'
' @param a The first value of the AND operation.
' @param b The second value of the AND operation.
' @return A new <a href="Int64Static.Int64.html">Int64</a> value containing the performed bitwise operation.
Public Function BitwiseAnd(ByRef a As Int64, ByRef b As Int64) As Int64
    BitwiseAnd.HighPart = a.HighPart And b.HighPart
    BitwiseAnd.LowPart = a.LowPart And b.LowPart
End Function

''
' Performs a bitwise OR of two <a href="Int64Static.Int64.html">Int64</a> values.
'
' @param a The first value of the OR operation.
' @param b The second value of the OR operation.
' @return A new <a href="Int64Static.Int64.html">Int64</a> value containing the performed bitwise operation.
Public Function BitwiseOr(ByRef a As Int64, ByRef b As Int64) As Int64
    BitwiseOr.HighPart = a.HighPart Or b.HighPart
    BitwiseOr.LowPart = a.LowPart Or b.LowPart
End Function

''
' Performs a bitwise NOT of two <a href="Int64Static.Int64.html">Int64</a> values.
'
' @param a The first value of the NOT operation.
' @param b The second value of the NOT operation.
' @return A new <a href="Int64Static.Int64.html">Int64</a> value containing the performed bitwise operation.
Public Function BitwiseNot(ByRef a As Int64) As Int64
    BitwiseNot.HighPart = Not a.HighPart
    BitwiseNot.LowPart = Not a.LowPart
End Function

''
' Performs a bitwise XOR of two <a href="Int64Static.Int64.html">Int64</a> values.
'
' @param a The first value of the XOR operation.
' @param b The second value of the XOR operation.
' @return A new <a href="Int64Static.Int64.html">Int64</a> value containing the performed bitwise operation.
Public Function BitwiseXor(ByRef a As Int64, ByRef b As Int64) As Int64
    BitwiseXor.HighPart = a.HighPart Xor b.HighPart
    BitwiseXor.LowPart = a.LowPart Xor b.LowPart
End Function

''
' Performs a left shift operation on an <a href="Int64Static.Int64.html">Int64</a>.
'
' @param Value The to shift left.
' @param BitCount The number of bits the value is to be shifted left.
' @return Returns a new <a href="Int64Static.Int64.html">Int64</a> containing the shifted value.
' @remarks If <i>BitCount</i> is negative, then a right shift is performed.
Public Function ShiftLeft(ByRef Value As Int64, ByVal BitCount As Long) As Int64
    If BitCount = 0 Then
        ShiftLeft = Value
    ElseIf BitCount < 0 Then
        ShiftLeft = ShiftRight(Value, -BitCount)
    ElseIf BitCount = 1 Then
        ShiftLeft.HighPart = ShiftLeftInt32(Value.HighPart, 1)
        
        If Value.LowPart And &H80000000 Then
            ShiftLeft.HighPart = ShiftLeft.HighPart Or 1
        End If
        
        ShiftLeft.LowPart = ShiftLeftInt32(Value.LowPart, 1)
    ElseIf BitCount < 64 Then
        If BitCount > 32 Then
            ShiftLeft.HighPart = ShiftLeftInt32(Value.LowPart, BitCount - 32)
        Else
            Dim Mask As Long
            Dim Bits As Long
        
            Mask = Not Powers(32 - BitCount) - 1
            Bits = Value.LowPart And Mask
            Bits = ShiftRightInt32(Bits, 32 - BitCount)
            ShiftLeft.HighPart = ShiftLeftInt32(Value.HighPart, BitCount)
            ShiftLeft.LowPart = ShiftLeftInt32(Value.LowPart, BitCount)
            ShiftLeft.HighPart = ShiftLeft.HighPart Or Bits
        End If
    End If
End Function

''
' Performs a right shift operation on an <a href="Int64Static.Int64.html">Int64</a>.
'
' @param Value The to shift right.
' @param BitCount The number of bits the value is to be shifted right.
' @return Returns a new <a href="Int64Static.Int64.html">Int64</a> containing the shifted value.
' @remarks If <i>BitCount</i> is negative, then a left shift is performed.
Public Function ShiftRight(ByRef Value As Int64, ByVal BitCount As Long) As Int64
    If BitCount = 0 Then
        ShiftRight = Value
    ElseIf BitCount < 0 Then
        ShiftRight = ShiftLeft(Value, -BitCount)
    ElseIf BitCount < 64 Then
        If BitCount > 32 Then
            ShiftRight.LowPart = ShiftRightInt32(Value.HighPart, BitCount - 32)
        Else
            Dim Mask As Long
            Dim Bits As Long
            
            Mask = Not Powers(BitCount)
            Bits = Value.HighPart And Mask
            Bits = ShiftLeftInt32(Bits, 32 - BitCount)
            ShiftRight.LowPart = ShiftRightInt32(Value.LowPart, BitCount)
            ShiftRight.HighPart = ShiftRightInt32(Value.HighPart, BitCount)
            ShiftRight.LowPart = ShiftRight.LowPart Or Bits
        End If
    End If
End Function

''
' Converts the string representation of a number to its 64-bit signed integer equivalent.
'
' @param s A string containing a number to convert.
' @return A 64-bit signed integer equivalent to the number contained in <i>s</i>.
' @remarks <p>The <i>s</i> parameter contains a number of the form:</p>
' <p>[ws][sign]digits[ws]</p>
' <p>Elements in square brackets ([ and ]) are optional. The following table describes each element.</p>
' <table style="width: auto">
'   <tr>
'       <th>Element</th>
'       <th>Description</th>
'   </tr>
'   <tr>
'        <td>ws</td><td>Optional white space.</td>
'   </tr>
'    <tr>
'        <td>sign</td><td>An optional sign.</td>
'    </tr>
'    <tr>
'        <td>digits</td><td>A sequence of digits ranging from 0 to 9.</td>
'    </tr>
' </table>
' <p>The <i>s</i> parameter is interpreted using the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.Integer</a> style. In addition to decimal digits, only leading and trailing spaces together with a leading sign are allowed. To explicitly define the style elements that can be present in <i>s</i>, use <a href="Int64Static.ParseEx.html">Int64.ParseEx(String, NumberStyles)</a> or the <a href="Int64Static.ParseEx.html">Int64.Parse(String, NumberStyles, IFormatProvider)</a> method.</p>
Public Function Parse(ByRef s As String) As Int64
    If Not TryParse(s, Parse) Then
        Error.Format Format_InvalidString
    End If
End Function

''
' Converts the string representation of a number in a specified style and culture-specific format to its 64-bit signed integer equivalent.
'
' @param s A string containing a number to convert.
' @param Style A bitwise combination of enumeration values that indicates the style elements that can be present in <i>s</i>. A typical value to specify is <a href="NumberFormatInfoStatic.NumberStyles.html">Integer</a>.
' @param Provider An <a href="IFormatProvider.html">IFormatProvider</a> that supplies culture-specific formatting information about <i>s</i>.
' @return A 64-bit signed integer equivalent to the number specified in <i>s</i>.
' @remarks <p>The <i>Style</i> parameter defines the style elements (such as white space or the positive sign) that are allowed in the <i>s</i> parameter for the parse operation to succeed. It must be a combination of bit flags from the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles</a> enumeration. Depending on the value of <i>Style</i>, the <i>s</i> parameter may include the following elements:</p>
' <p>[ws][$][sign][digits,]digits[.fractional_digits][e[sign]exponential_digits][ws]</p>
' <p>Or, if style includes <a href="NumberFormatInfoStatic.NumberStyles.html">AllowHexSpecifier</a>:</p>
' <p>[ws]hexdigits[ws]</p>
' <p>Elements in square brackets ([ and ]) are optional. The following table describes each element.</p>
' <table style="width: auto">
'   <tr>
'       <th>Element</th>
'       <th>Description</th>
'   </tr>
'   <tr>
'        <td><i>ws</i></td><td>Optional white space. White space can appear at the beginning of <i>s</i> if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowLeadingWhite</a> flag, and it can appear at the end of <i>s</i> if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowTrailingWhite</a> flag.</td>
'   </tr>
'    <tr>
'        <td><i>$</i></td><td>A culture-specific currency symbol. Its position in the string is defined by the <a href="NumberFormatInfo.get.CurrencyPositivePattern.html">NumberFormatInfo.CurrencyPositivePattern</a> property of the <a href="NumberFormatInfo.html">NumberFormatInfo</a> object returned by the <a href="IFormatProvider.GetFormat.html">GetFormat</a> method of the <i>Provider</i> parameter. The currency symbol can appear in <i>s</i> if style includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowCurrencySymbol</a> flag.</td>
'    </tr>
'    <tr>
'        <td>
'           <i>digits</i><br/>
'           <i>fractional_digits</i><br/>
'           <i>expontial_digits</i>
'        </td>
'        <td>A sequence of digits from 0 through 9.</td>
'    </tr>
'    <tr>
'        <td><i>,</i></td><td>A culture-specific thousands separator symbol. The thousands separator of the culture specified by <i>Provider</i> can appear in <i>s</i> if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowThousands</a> flag.</td>
'    </tr>
'    <tr>
'        <td><i>.</i></td>
'        <td>
'           A culture-specific decimal point symbol. The decimal point symbol of the culture specified by <i>Provider</i> can appear in <i>s</i> if style includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowDecimalPoint</a> flag.<br/>
'           Only the digit 0 can appear as a fractional digit for the parse operation to succeed; if <i>fractional_digits</i> includes any other digit, an <a href="OverflowException.html">OverflowException</a> is thrown.
'        </td>
'    </tr>
'    <tr>
'        <td><i>e</i></td><td>The 'e' or 'E' character, which indicates that the value is represented in exponential notation. The <i>s</i> parameter can represent a number in exponential notation if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowExponent</a> flag.</td>
'    </tr>
'    <tr>
'        <td><i>hexdigits</i></td><td>A sequence of hexadecimal digits from 0 through f, or 0 through F.</td>
'    </tr>
' </table>
Public Function ParseEx(ByRef s As String, ByVal Style As NumberStyles, ByVal Provider As IFormatProvider) As Int64
    If Not TryParseEx(s, Style, Provider, ParseEx) Then
        Error.Format Format_InvalidString
    End If
End Function

''
' Converts the string representation of a number to its 64-bit signed integer equivalent. A return value indicates whether the conversion succeeded or failed.
'
' @param s A string containing a number to convert.
' @param Result When this method returns, contains the 64-bit signed integer value equivalent of the number contained in <i>s</i>, if the conversion succeeded, or zero if the conversion failed. The conversion fails if the <i>s</i> parameter is empty, is not in a format compliant withstyle, or represents a number less than <a href="Int64Static.get.MinValue.html">MinValue</a> or greater than <a href="Int64Static.get.MaxValue.html">MaxValue</a>. This parameter is passed uninitialized; any value originally supplied in <i>Result</i> will be overwritten.
' @return <b>True</b> if s was converted successfully; otherwise, <b>False</b>.
' @remarks <p>The <b>TryParse</b> method is like the <a href="Int64Static.Parse.html">Parse</a> method, except the <b>TryParse</b> method does not throw an exception if the conversion fails. It eliminates the need to use exception handling to test for a <a href="FormatException.html">FormatException<a/> in the event that <i>s</i> is invalid and cannot be successfully parsed.</p>
' <p>The s parameter contains a number of the form:</p>
' <p>[ws][sign]digits[ws]</p>
' <p>Elements in square brackets ([ and ]) are optional. The following table describes each element.</p>
' <table style="width: auto">
'   <tr>
'       <th>Element</th>
'       <th>Description</th>
'   </tr>
'   <tr>
'        <td>ws</td><td>Optional white space.</td>
'   </tr>
'    <tr>
'        <td>sign</td><td>An optional sign.</td>
'    </tr>
'    <tr>
'        <td>digits</td><td>A sequence of digits ranging from 0 to 9.</td>
'    </tr>
' </table>
' <p>The <i>s</i> parameter is interpreted using the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.Integer</a> style. In addition to decimal digits, only leading and trailing spaces together with a leading sign are allowed. To explicitly define the style elements that can be present in <i>s</i>, use <a href="Int64Static.ParseEx.html">Int64.ParseEx(String, NumberStyles)</a> or the <a href="Int64Static.ParseEx.html">Int64.Parse(String, NumberStyles, IFormatProvider)</a> method.</p>
Public Function TryParse(ByRef s As String, ByRef Result As Int64) As Boolean
    TryParse = TryParseEx(s, NumberStyles.Integer, Nothing, Result)
End Function

''
' Converts the string representation of a number in a specified style and culture-specific format to its 64-bit signed integer equivalent.
'
' @param s A string containing a number to convert.
' @param Style A bitwise combination of enumeration values that indicates the style elements that can be present in <i>s</i>. A typical value to specify is <a href="NumberFormatInfoStatic.NumberStyles.html">Integer</a>.
' @param Provider An <a href="IFormatProvider.html">IFormatProvider</a> that supplies culture-specific formatting information about <i>s</i>.
' @param Result When this method returns, contains the 64-bit signed integer value equivalent of the number contained in <i>s</i>, if the conversion succeeded, or zero if the conversion failed. The conversion fails if the <i>s</i> parameter is empty, is not in a format compliant withstyle, or represents a number less than <a href="Int64Static.get.MinValue.html">MinValue</a> or greater than <a href="Int64Static.get.MaxValue.html">MaxValue</a>. This parameter is passed uninitialized; any value originally supplied in <i>Result</i> will be overwritten.
' @return <b>True</b> if s was converted successfully; otherwise, <b>False</b>.
' @remarks <p>The <i>Style</i> parameter defines the style elements (such as white space or the positive sign) that are allowed in the <i>s</i> parameter for the parse operation to succeed. It must be a combination of bit flags from the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles</a> enumeration. Depending on the value of <i>Style</i>, the <i>s</i> parameter may include the following elements:</p>
' <p>[ws][$][sign][digits,]digits[.fractional_digits][e[sign]exponential_digits][ws]</p>
' <p>Or, if style includes <a href="NumberFormatInfoStatic.NumberStyles.html">AllowHexSpecifier</a>:</p>
' <p>[ws]hexdigits[ws]</p>
' <p>Elements in square brackets ([ and ]) are optional. The following table describes each element.</p>
' <table style="width: auto">
'   <tr>
'       <th>Element</th>
'       <th>Description</th>
'   </tr>
'   <tr>
'        <td><i>ws</i></td><td>Optional white space. White space can appear at the beginning of <i>s</i> if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowLeadingWhite</a> flag, and it can appear at the end of <i>s</i> if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowTrailingWhite</a> flag.</td>
'   </tr>
'    <tr>
'        <td><i>$</i></td><td>A culture-specific currency symbol. Its position in the string is defined by the <a href="NumberFormatInfo.get.CurrencyPositivePattern.html">NumberFormatInfo.CurrencyPositivePattern</a> property of the <a href="NumberFormatInfo.html">NumberFormatInfo</a> object returned by the <a href="IFormatProvider.GetFormat.html">GetFormat</a> method of the <i>Provider</i> parameter. The currency symbol can appear in <i>s</i> if style includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowCurrencySymbol</a> flag.</td>
'    </tr>
'    <tr>
'        <td>
'           <i>digits</i><br/>
'           <i>fractional_digits</i><br/>
'           <i>expontial_digits</i>
'        </td>
'        <td>A sequence of digits from 0 through 9.</td>
'    </tr>
'    <tr>
'        <td><i>,</i></td><td>A culture-specific thousands separator symbol. The thousands separator of the culture specified by <i>Provider</i> can appear in <i>s</i> if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowThousands</a> flag.</td>
'    </tr>
'    <tr>
'        <td><i>.</i></td>
'        <td>
'           A culture-specific decimal point symbol. The decimal point symbol of the culture specified by <i>Provider</i> can appear in <i>s</i> if style includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowDecimalPoint</a> flag.<br/>
'           Only the digit 0 can appear as a fractional digit for the parse operation to succeed; if <i>fractional_digits</i> includes any other digit, an <a href="OverflowException.html">OverflowException</a> is thrown.
'        </td>
'    </tr>
'    <tr>
'        <td><i>e</i></td><td>The 'e' or 'E' character, which indicates that the value is represented in exponential notation. The <i>s</i> parameter can represent a number in exponential notation if <i>Style</i> includes the <a href="NumberFormatInfoStatic.NumberStyles.html">NumberStyles.AllowExponent</a> flag.</td>
'    </tr>
'    <tr>
'        <td><i>hexdigits</i></td><td>A sequence of hexadecimal digits from 0 through f, or 0 through F.</td>
'    </tr>
' </table>
Public Function TryParseEx(ByRef s As String, ByVal Style As NumberStyles, ByVal Provider As IFormatProvider, ByRef Result As Int64) As Boolean
    Dim Info As NumberFormatInfo
    
    Result.LowPart = 0
    Result.HighPart = 0
    
    If IsHexSpecifier(Style) Then
        TryParseEx = TryParseHex(s, Style, Result)
    Else
        Set Info = NumberFormatInfo.GetInstance(Provider)
        TryParseEx = TryParseDecimal(s, Style, Info, Result)
    End If
    
    If Not TryParseEx Then
        Result.LowPart = 0
        Result.HighPart = 0
    End If
End Function

Private Function IsHexSpecifier(ByVal Style As NumberStyles) As Boolean
    IsHexSpecifier = (Style And NumberStyles.AllowHexSpecifier)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MultiplyByNegative(ByRef a As Int64, ByRef Negative As Int64, ByRef Result As Int64)
    MultiplyPositives a, Negate(Negative), Result
    Result = Negate(Result)
End Sub

Private Sub MultiplyPositives(ByRef a As Int64, ByRef b As Int64, ByRef Result As Int64)
    Dim u As QWordArray
    Dim v As QWordArray
    Dim r As QWordArray
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim d As Long
        
    u = AsQWordArr(a)
    v = AsQWordArr(b)
    
    For i = 0 To 3
        k = 0
        d = v.Words(i) And &HFFFF&
                
        For j = 0 To 3
#If Release Then
            k = d * (u.Words(j) And &HFFFF&) + (r.Words(j) And &HFFFF&) + k
            r.Words(j) = k And &HFFFF&
            k = ((k And &HFFFF0000) \ vbShift16Bits) And &HFFFF&
#Else
            k = UInt16x16To32(d, u.Words(j)) + (r.Words(j) And &HFFFF&) + k
            r.Words(j) = AsWord(k)
            k = ((k And &HFFFF0000) \ vbShift16Bits) And &HFFFF&
#End If
        Next
    Next
    
    AssignQWord r, Result
End Sub

#If Not Release Then
Private Function UInt16x16To32(ByVal x As Long, ByVal y As Long) As Long
    Dim u As Currency
    Dim v As Currency
    Dim w As Currency
    
    u = x And &HFFFF&
    v = y And &HFFFF&
    w = (u * v) * 0.0001@
      
    UInt16x16To32 = AsLong(w)
End Function
#End If

' adapted from https://en.wikipedia.org/wiki/Division_algorithm
'
' Integer division (unsigned) with remainder
'
' Q := 0                  -- Initialize quotient and remainder to zero
' R := 0
' for i := n - 1 .. 0 do  -- Where n is number of bits in N
'   R := R << 1           -- Left-shift R by 1 bit
'   R(0) := N(i)          -- Set the least-significant bit of R equal to bit i of the numerator
'   If R >= D Then
'     R := R - D
'     Q(i) := 1
'   End
' End
Private Sub DivideCore(ByRef Dividend As Int64, ByRef Divisor As Int64, ByRef Quotient As Int64, ByRef Remainder As Int64)
    Dim n As Long
    Dim i As Long
    
    Debug.Assert (Dividend.HighPart And &H80000000) = 0
    Debug.Assert (Divisor.HighPart And &H80000000) = 0
    
    For i = MaxBit(Dividend) - 1 To 0 Step -1
        InPlaceShiftLeftOne Remainder
        
        If i > 31 Then
            n = Dividend.HighPart And Powers(i - 32)
        Else
            n = Dividend.LowPart And Powers(i)
        End If
        
        If n Then
            Remainder.LowPart = Remainder.LowPart Or 1
        End If
        
        If Compare(Remainder, Divisor) >= 0 Then
            InPlaceSubtract Remainder, Divisor
            
            If i > 31 Then
                Quotient.HighPart = Quotient.HighPart Or Powers(i - 32)
            Else
                Quotient.LowPart = Quotient.LowPart Or Powers(i)
            End If
        End If
    Next
End Sub

Private Sub InPlaceSubtract(ByRef a As Int64, ByRef b As Int64)
    AsCurr(a) = AsCurr(a) - AsCurr(b)
End Sub

Private Sub InPlaceShiftLeftOne(ByRef Value As Int64)
    Value.HighPart = ShiftLeftOne(Value.HighPart)
    
    If Value.LowPart And &H80000000 Then
        Value.HighPart = Value.HighPart Or 1
    End If
    
    Value.LowPart = ShiftLeftOne(Value.LowPart)
End Sub

Private Function ShiftLeftOne(ByVal Value As Long) As Long
    Dim HighBit As Long
    
    If Value And &H40000000 Then
        HighBit = &H80000000
    End If
    
    ShiftLeftOne = ((Value And &H3FFFFFFF) * vbShift1Bit) Or HighBit
End Function

Private Function MaxBit(ByRef Value As Int64) As Long
    If Value.HighPart Then
        MaxBit = MaxLongBit(Value.HighPart) + 32
    ElseIf Value.LowPart Then
        MaxBit = MaxLongBit(Value.LowPart)
    End If
End Function

Private Function MaxLongBit(ByVal Value As Long) As Long
    If Value And &H80000000 Then
        MaxLongBit = 32
    ElseIf Value And &HFF000000 Then
        MaxLongBit = MaxByteBit((Value And &HFF000000) \ vbShift24Bits) + 24
    ElseIf Value And &HFF0000 Then
        MaxLongBit = MaxByteBit((Value And &HFF0000) \ vbShift16Bits) + 16
    ElseIf Value And &HFF00& Then
        MaxLongBit = MaxByteBit((Value And &HFF00&) \ vbShift8Bits) + 8
    ElseIf Value And &HFF& Then
        MaxLongBit = MaxByteBit(Value And &HFF&)
    End If
End Function

Private Function MaxByteBit(ByVal Value As Byte) As Long
    If Value And &H80 Then
        MaxByteBit = 8
    ElseIf Value And &H40 Then
        MaxByteBit = 7
    ElseIf Value And &H20 Then
        MaxByteBit = 6
    ElseIf Value And &H10 Then
        MaxByteBit = 5
    ElseIf Value And &H8 Then
        MaxByteBit = 4
    ElseIf Value And &H4 Then
        MaxByteBit = 3
    ElseIf Value And &H2 Then
        MaxByteBit = 2
    ElseIf Value And &H1 Then
        MaxByteBit = 1
    End If
End Function

Private Function TryParseDecimal(ByRef s As String, ByVal Style As NumberStyles, ByVal Info As NumberFormatInfo, ByRef n As Int64) As Boolean
    Const NotFound          As Long = 0
    Const FoundBefore       As Long = 1
    Const FoundAfter        As Long = 2
    Dim Chars()             As Integer
    Dim Index               As Long
    Dim MaxIndex            As Long
    Dim Ch                  As Long
    Dim InWhitespace        As Boolean
    Dim NeedsClosingParen   As Boolean
    Dim NegativeSign        As Integer
    Dim PositiveSign        As Integer
    Dim CurrencySymbol      As Integer
    Dim Thousands           As Integer
    Dim CurrencyThousands   As Integer
    Dim DecimalPoint        As Integer
    Dim CurrencyPoint       As Integer
    Dim FoundDigit          As Boolean
    Dim FoundCurrency       As Long
    Dim FoundSign           As Long
    Dim FoundDecimal        As Boolean
    Dim FoundClosingParen   As Boolean
    Dim FoundExponent       As Boolean
    Dim Exponent            As Long
    Dim DecimalPlaces       As Long
    Dim Ten                 As Int64
    Dim Char64              As Int64
    Dim IsNegative          As Boolean
    Dim IsMaxValue          As Boolean
    
    Ten.LowPart = 10
    NegativeSign = AscW(Info.NegativeSign)
    PositiveSign = AscW(Info.PositiveSign)
    CurrencySymbol = AscW(Info.CurrencySymbol)
    DecimalPoint = AscW(Info.NumberDecimalSeparator)
    CurrencyPoint = AscW(Info.CurrencyDecimalSeparator)
    Thousands = AscW(Info.NumberGroupSeparator)
    CurrencyThousands = AscW(Info.CurrencyGroupSeparator)
    Chars = AllocChars(s)
    MaxIndex = UBound(Chars)
    
    On Error GoTo Catch
    
    Do While Index <= MaxIndex
        Ch = Chars(Index)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                If IsMaxValue Then _
                    Throw New OverflowException
                
                If (FoundCurrency = FoundAfter) Or (FoundSign = FoundAfter) Or InWhitespace Then
                    GoTo Finish
                End If
                
                If Not FoundDecimal Then
                    n = Multiply(n, Ten)
                    
                    If IsNegative Then
                        If n.HighPart = &H7FFFFFFF Then
                            If n.LowPart = &HFFFFFFF8 Then
                                Select Case Ch
                                    Case vbEightChar
                                        IsMaxValue = True
                                        GoTo Continue
                                    Case vbNineChar
                                        Throw New OverflowException
                                End Select
                            End If
                        End If
                    End If
                    
                    Char64.LowPart = Ch - vbZeroChar
                    n = Add(n, Char64)
                Else
                    If Style And NumberStyles.AllowExponent Then
                        ' we are allowing a possible exponent value to be parsed,
                        ' so if we have encountered a decimal point, we need to
                        ' continue parsing digits and also count how many
                        ' decimal places we have gone.
                        n = Multiply(n, Ten)
                        Char64.LowPart = Ch - vbZeroChar
                        n = Add(n, Char64)
                        DecimalPlaces = DecimalPlaces + 1
                    Else
                        If Ch <> vbZeroChar Then
                            GoTo Finish
                        End If
                    End If
                End If
                
                FoundDigit = True
            Case NegativeSign, PositiveSign
                If (FoundSign <> NotFound) Or NeedsClosingParen Then
                    GoTo Finish
                End If
                
                If Not FoundDigit Then
                    If (Style And NumberStyles.AllowLeadingSign) = 0 Then
                        GoTo Finish
                    End If
                    
                    FoundSign = FoundBefore
                Else
                    If (Style And NumberStyles.AllowTrailingSign) = 0 Then
                        GoTo Finish
                    End If
                    
                    FoundSign = FoundAfter
                End If
                
                If Ch = NegativeSign Then
                    IsNegative = True
                End If
            Case vbTabChar, vbLfChar, vbVtChar, vbFfChar, vbCrChar, vbSpaceChar
                If Not FoundDigit Then
                    If ((Style And NumberStyles.AllowLeadingWhite) = 0) Or NeedsClosingParen Then
                        GoTo Finish
                    End If
                
                    InWhitespace = FoundSign
                Else
                    InWhitespace = True
                End If
            Case vbOpenParenChar
                If ((Style And NumberStyles.AllowParentheses) = 0) Or (FoundSign <> NotFound) Then
                    GoTo Finish
                End If
                
                NeedsClosingParen = True
            Case vbCloseParenChar
                If (Style And NumberStyles.AllowParentheses) = 0 Or (Not NeedsClosingParen) Then
                    GoTo Finish
                End If
                
                IsNegative = True
                NeedsClosingParen = False
                FoundClosingParen = True
            Case CurrencySymbol
                If ((Style And NumberStyles.AllowCurrencySymbol) = 0) Or (FoundCurrency <> NotFound) Then
                    GoTo Finish
                End If
                
                FoundCurrency = IIfLong(IsZero(n), FoundBefore, FoundAfter)
            Case DecimalPoint, CurrencyPoint
                If ((Style And NumberStyles.AllowDecimalPoint) = 0) Or FoundDecimal Or FoundClosingParen Then
                    GoTo Finish
                End If
                
                If FoundCurrency Then
                    If Ch = DecimalPoint Then
                        GoTo Finish
                    End If
                End If
                
                FoundDecimal = True
            Case Thousands, CurrencyThousands
                If ((Style And NumberStyles.AllowThousands) = 0) Or (Not FoundDigit) Then
                    GoTo Finish
                End If
                
                If FoundCurrency Then
                    If Ch = Thousands Then
                        GoTo Finish
                    End If
                End If
            Case vbLowerEChar, vbUpperEChar
                If ((Style And NumberStyles.AllowExponent) = 0) Or FoundExponent Then
                    GoTo Finish
                End If
            
                If Not TryParseExponent(Chars, Index, Exponent) Then
                    GoTo Finish
                End If
                
                ' we understand that parsing the exponent will advance the index passed the end of the digits
                ' and that we will be advancing the index as well, in effect advancing it twice, so we need
                ' to counter that last advancement.
                Index = Index - 1
                FoundExponent = True
            Case Else
                GoTo Finish
        End Select
        
Continue:
        Index = Index + 1
    Loop
    
    If InWhitespace Then
        If (Style And NumberStyles.AllowTrailingWhite) = 0 Then
            GoTo Finish
        End If
    End If
    
    If NeedsClosingParen Then
        GoTo Finish
    End If
    
    If FoundExponent Then
        If Exponent < DecimalPlaces Then
            GoTo Finish
        End If
    
        Do While Exponent - DecimalPlaces > 0
            n = Multiply(n, Ten)
            Exponent = Exponent - 1
        Loop
    End If
    
    If IsNegative Then
        If IsMaxValue Then
            n.HighPart = &H80000000
            n.LowPart = 0
        Else
            n = Negate(n)
        End If
    End If
    
    TryParseDecimal = True
    
Finish:
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars Chars
    Return
End Function

Private Function TryParseExponent(ByRef Chars() As Integer, ByRef Index As Long, ByRef Exponent As Long) As Boolean
    Dim Ch          As Integer
    Dim MaxIndex    As Long
    
    MaxIndex = UBound(Chars)
    Index = Index + 1
    
    If Index > MaxIndex Then
        Exit Function
    End If
    
    If Chars(Index) = vbPlusChar Then
        Index = Index + 1
    End If
    
    Do While Index <= MaxIndex
        Ch = Chars(Index)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                Exponent = Exponent * 10 + (Ch - vbZeroChar)
                TryParseExponent = True
            Case Else
                Exit Function
        End Select
        
        Index = Index + 1
    Loop
End Function

Private Function TryParseHex(ByRef s As String, ByVal Style As NumberStyles, ByRef n As Int64) As Boolean
    Const vbLowerHexOffset As Long = 87
    Const vbUpperHexOffset As Long = 55
    
    Dim Chars()     As Integer
    Dim FoundDigits As Boolean
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Ch As Long
    Dim Section As Long
    
    Chars = AllocChars(s)
    j = 1
    Section = 1
    
    For i = UBound(Chars) To LBound(Chars) Step -1
        Select Case Chars(i)
            Case vbZeroChar
                Ch = 0
                FoundDigits = True
            Case vbOneChar To vbNineChar
                Ch = Chars(i) - vbZeroChar
                FoundDigits = True
            Case vbLowerAChar To vbLowerFChar
                Ch = Chars(i) - vbLowerHexOffset
                FoundDigits = True
            Case vbUpperAChar To vbUpperFChar
                Ch = Chars(i) - vbUpperHexOffset
                FoundDigits = True
            Case vbTabChar, vbLfChar, vbVtChar, vbFfChar, vbCrChar, vbSpaceChar
                If ((Style And NumberStyles.AllowLeadingWhite) = 0) And FoundDigits Then
                    GoTo Finally
                End If

                If ((Style And NumberStyles.AllowTrailingWhite) = 0) And Not FoundDigits Then
                    GoTo Finally
                End If

                GoTo Continue
            Case Else
                GoTo Finally
        End Select

        k = k Or (Ch * j)
        j = j * vbShift4Bits

        If j = vbShift16Bits Then
            Select Case Section
                Case 1
                    n.LowPart = k
                Case 2
                    n.LowPart = n.LowPart Or ShiftLeftInt32(k, 16)
                Case 3
                    n.HighPart = k
                Case 4
                    n.HighPart = n.HighPart Or ShiftLeftInt32(k, 16)
                Case Else
                    GoTo Finally
            End Select
            
            Section = Section + 1
            j = 1
            k = 0
        End If

Continue:
    Next i

    If k Then
        Select Case Section
            Case 1
                n.LowPart = k
            Case 2
                n.LowPart = n.LowPart Or ShiftLeftInt32(k, 16)
            Case 3
                n.HighPart = k
            Case 4
                n.HighPart = n.HighPart Or ShiftLeftInt32(k, 16)
            Case Else
                GoTo Finally
        End Select
    End If
    TryParseHex = True

Finally:
    FreeChars Chars
End Function
