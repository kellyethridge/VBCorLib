VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigIntegerStatic"
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
' Module: BigIntegerStatic
'

''
' Provides methods used to manipulate <a href="BigInteger.html">BigInteger</a> objects.
'
' @remarks The methods provided by this class are accessed through the convention of <b>BigInteger.</i>MethodName</i></b>.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Rnd(128)
' Set b = BigInteger.Pow(b, BInt(3))
' </pre>
' <p>Most of these methods exist in the <a href="BigInteger.html">BigInteger</a> class and can be used
' directly with <a href="BigInteger.html">BigInteger</a> objects. The primary advantage of using the
' methods defined in this class is that a value of <b>Nothing</b> can be passed in and
' will be treated as a value of 0.</p>
Option Explicit

Private mZero       As BigInteger
Private mOne        As BigInteger
Private mMinusOne   As BigInteger
Private mTen        As BigInteger


''
' Returns a shared <a href="BigInteger.html">BigInteger</a> instance with the value of 0.
'
' @return Returns a shared instance of the value of 0.
' @remarks This is to reduce the creation of multiple instances of values that are commonly used.
Public Property Get Zero() As BigInteger
    Set Zero = mZero
End Property

''
' Returns a shared <a href="BigInteger.html">BigInteger</a> instance with the value of 1.
'
' @return Returns a shared instance of the value of 1.
' @remarks This is to reduce the creation of multiple instances of values that are commonly used.
Public Property Get One() As BigInteger
    Set One = mOne
End Property

''
' Returns a shared <a href="BigInteger.html">BigInteger</a> instance with the value of -1.
'
' @return Returns a shared instance of the value of -1.
' @remarks This is to reduce the creation of multiple instances of values that are commonly used.
Public Property Get MinusOne() As BigInteger
    Set MinusOne = mMinusOne
End Property

''
' Returns a shared <a href="BigInteger.html">BigInteger</a> instance with the value of 10.
'
' @return Returns a shared instance of the value of 10.
' @remarks This is to reduce the creation of multiple instances of values that are commonly used.
Public Property Get Ten() As BigInteger
    Set Ten = mTen
End Property

''
' Compares two <a href="BigInteger.html">BigInteger</a> instances to each other returning their relation.
'
' @param x The left handed value to be compared to.
' @param y The right handed value to be compared to.
' @return Returns the relation between to the two numbers.<br>
' @remarks If <b>Nothing</b> is passed in, it is treated as zero.<br>
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>Left</i> is less than <i>Right</i>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>Left<i/> equals <i>Right</i>.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>Left</i> is greater than <i>Right</i>.</td>
'    </tr>
' </table>
' <pre>
' Dim a As BigInteger, b As BigInteger, Relation As String
' Set a = BInt(1)
' Set b = BInt(2)
'
' Select Case BigInteger.Compare(a, b)
'     Case Is < 0: Relation = "<"
'     Case Is > 0: Relation = ">"
'     Case Else: Relation = "="
' End Select
'
' Console.WriteLine "{0} {1} {2}", a, Relation, b
' </pre>
Public Function Compare(ByVal Left As BigInteger, ByVal Right As BigInteger) As Long
    Set Left = ValueOrZero(Left)
    Compare = Left.CompareTo(Right)
End Function

''
' Returns the equality of two <a href="BigInteger.html">BigInteger</a> values.
'
' @param x The left handed value to test for equality.
' @param y The right handed value to test for equality.
' @return <b>True</b> if the values are equal, <b>False</b> otherwise.
' @remarks If <b>Nothing</b> is passed in, it is treated as zero.
Public Function Equals(ByVal x As BigInteger, ByVal y As BigInteger) As Boolean
    Set x = ValueOrZero(x)
    Equals = x.Equals(y)
End Function

''
' Attemps to parse a string of characters of the specified base type.
'
' @param Value The string of characters to parse.
' @param OutResult The returning <a href="BigInteger.html">BigInteger</a> object if the string was successfully parsed.
' @return <b>True</b> if the string was parsed, <b>False</b> otherwise.
' @remarks 2 base types are supported for parsing; Decimal and Hex. The string
' must include a base type specifier for Hex. Decimal is the default.
' <p>Hex Prefix Specifiers: 0x, 0X, &h, &H</p>
' <p>Parsing is not case sensitive.</p>
Public Function TryParse(ByRef Value As String, ByRef OutResult As BigInteger) As Boolean
    TryParse = TryParseEx(Value, NumberStyles.Integer, Nothing, OutResult)
End Function

''
' Tries to convert the string representation of a number in a specified style and culture-specific format to its <a href="BigInteger.html">BigInteger</a> equivalent, and returns a value that indicates whether the conversion succeeded.
'
' @param Value The string representation of a number. The string is interpreted using the style specified by style.
' @param Style A bitwise combination of enumeration values that indicates the style elements that can be present in value. A typical value to specify is NumberStyles.Integer.
' @param Provider An object that supplies culture-specific formatting information about value. If Nothing is passed in the current CultureInfo will be used.
' @param OutResult When this method returns, contains the <a href="BigInteger.html">BigInteger</a> equivalent to the number that is contained in value, or <a href="BigInteger.Get.Zero.html">BigInteger.Zero</a> if the conversion failed. The conversion fails if thevalue parameter is null or is not in a format that is compliant withstyle. This parameter is passed uninitialized.
Public Function TryParseEx(ByRef Value As String, ByVal Style As NumberStyles, ByVal Provider As IFormatProvider, ByRef OutResult As BigInteger) As Boolean
    Dim Result      As BigNumber
    Dim Info        As NumberFormatInfo
    Dim HasPrefix   As Boolean
    
    Set OutResult = Nothing
    
    If IsHexSpecifier(Value, Style, HasPrefix) Then
        TryParseEx = TryParseHex(Value, HasPrefix, Style, Result)
    Else
        Set Info = NumberFormatInfo.GetInstance(Provider)
        TryParseEx = TryParseDecimal(Value, Style, Info, Result)
    End If
    
    If TryParseEx Then
        Set OutResult = New BigInteger
        OutResult.InitBigNumber Result
    End If
End Function

Private Function IsHexSpecifier(ByRef Value As String, ByVal Style As NumberStyles, ByRef HasPrefix As Boolean) As Boolean
    Select Case Left$(Value, 2)
        Case "&h", "&H", "0x", "0X"
            HasPrefix = True
            IsHexSpecifier = True
        Case Else
            IsHexSpecifier = (Style And NumberStyles.AllowHexSpecifier)
    End Select
End Function

''
' Attemps to parse a string of characters of the specified base type.
'
' @param Value The string of characters to attempt to be parsed.
' @param Style A bitwise combination of the enumeration values that specify the permitted format of <i>Value</i>.
' @param Provider An object that provides culture-specific formatting information about <i>Value</i>.
' @return Returns an instance of <a href="BigInteger.html">BigInteger</a> containing the parsed value.
' @remarks If the string could not be parsed, an exception is thrown.
' <p>3 base types are supported for parsing (Decimal, Hex.) The string
' must include a base type specifier for Hex. Decimal is the default.</p>
' <p>
' Hex Specifiers: 0x, 0X, &h, &H - (0xFF, 0XFF, &hFF, &HFF)
' </p>
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Parse("&H7FFF")
' Debug.Print b.ToString("X") '' 7FFF
' </pre>
' <p>Parsing is not case sensitive.</p>
Public Function Parse(ByRef Value As String, Optional ByVal Style As NumberStyles = NumberStyles.Integer, Optional ByVal Provider As IFormatProvider) As BigInteger
    If Not TryParseEx(Value, Style, Provider, Parse) Then
        Error.Format Format_ParseBigInteger
    End If
End Function

''
' Returns a human readable string of characters representing the number.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> to create a string from.
' @param Format The specific numerical format wish to have the number displayed as.
' @return The numeric value of the number in the specified format.
' @remarks "X" and "x" for hex in either upper or lowercase. Everything else defaults to a decimal display.
' <p>The specifier can also contain a number of mininum digits to be displayed. If
' the value isn't long enough to reach the mininum digits, then leading 0's or
' FF's will be added to the beginning of the output.</p>
' <pre>
' Dim b As BigInteger
' Set b = BInt(-200)
' Debug.Print b.ToString("d8") '' -00000200
' Debug.Print b.ToString("d")  '' -200
' Debug.Print b.ToString("x8") '' ffffff38
' Debug.Print b.ToString("x")  '' ff38
' </pre>
' <p>If no format is specified, the default of "d" is used.</p>
Public Function ToString(ByVal Value As BigInteger, Optional ByRef Format As String) As String
    Set Value = ValueOrZero(Value)
    ToString = Value.ToString(Format)
End Function

''
' Returns the opposite signed value of this <a href="BigInteger.html">BigInteger</a> instance.
'
' @param Value The value to be negated.
' @return Returns a <a href="BigInteger.html">BigInteger</a> instance containing the opposite signed value.
' @remarks Negative values are represented as a twos-complement number, unlike
' other systems used to represent a negative value. Any negation requires
' the twos-complement algorithm to be applied to the number. This keeps in step
' with how users of VB already deal with negative numbers, helping to ease
' the transition to using BigInteger values.
' <p>If <i>Value</i> is <b>Nothing</b> it will be treated as zero.</p>
Public Function Negate(ByVal Value As BigInteger) As BigInteger
    Set Value = ValueOrZero(Value)
    Set Negate = Value.Negate
End Function

''
' Returns the absolute (positive or zero) value of this instance.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> value to return the absolute value of.
' @return Returns a positive (or zero) value of the current <a href="BigInteger.html">BigInteger</a> instance.
' @remarks If the value is already positive or zero, then the same object instance
' is returned. If the value is negative, then the negated value is returned.
' <p><i>Value</i> is <b>Nothing</b> it will be treated as zero.</p>
Public Function Absolute(ByVal Value As BigInteger) As BigInteger
    Set Value = ValueOrZero(Value)
    Set Absolute = Value.Absolute
End Function

''
' Adds two <a href="BigInteger.html">BigInteger</a> value instances together, returning the sum.
'
' @param Left The first value to add.
' @param Right The second value to add.
' @return The sum of <i>Left</i> and <i>Right</i>.
' @remarks Values of <b>Nothing</b> will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Add(BInt(100), BInt(150))
' Debug.Print b.ToString '' 250
' </pre>
Public Function Add(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set Add = Left.Add(Right)
End Function

''
' Subtracts one <a href="BigInteger.html">BigInteger</a> value from another, returning the difference.
'
' @param Left The minuend value to be subracted from.
' @param Right The subtrahend value to subtract from the minuend value.
' @return The difference of the two values.
' @remarks Values of <b>Nothing</b> will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Subtract(BInt(100), BInt(150))
' Debug.Print b.ToString '' -50
' </pre>
Public Function Subtract(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set Subtract = Left.Subtract(Right)
End Function

''
' Multiplies <a href="BigInteger.html">BigInteger</a> values, returning their product.
'
' @param Left The multiplicand to be multiplied.
' @param Right The multiplier to multiply by.
' @return The product of the two values.
' @remarks Values of <b>Nothing</b> will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Multiply(BInt(100), BInt(150))
' Debug.Print b.ToString '' 15000
' </pre>
Public Function Multiply(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set Multiply = Left.Multiply(Right)
End Function


''
' Divides one <a href="BigInteger.html">BigInteger</a> value by another, returning the quotient.
'
' @param dividend The <a href="BigInteger.html">BigInteger</a> value to be divided.
' @param divisor The <a href="BigInteger.html">BigInteger</a> value to divide by.
' @return The quotient of the division.
' @remarks Values of <b>Nothing</b> will be treated as zero.
' <p>Integer division is applied, there is no rounding values. This method discards the remainder.</p>
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Divide(BInt(100), BInt(40))
' Debug.Print b.ToString '' 2
' </pre>
'
Public Function Divide(ByVal Dividend As BigInteger, ByVal Divisor As BigInteger) As BigInteger
    Set Dividend = ValueOrZero(Dividend)
    Set Divide = Dividend.Divide(Divisor)
End Function

''
' Divides one <a href="BigInteger.html">BigInteger</a> value by another, returning both the quotient and remainder.
'
' @param dividend The <a href="BigInteger.html">BigInteger</a> value to be divided.
' @param divisor The <a href="BigInteger.html">BigInteger</a> value to divide by.
' @param remainder A variable to place the remainder in to.
' @return The quotient of the division.
' @remarks Values of <b>Nothing</b> will be treated as zero.
' <pre>
' Dim b As BigInteger
' Dim r As BigInteger
' Dim q As BigInteger
' Set b = BInt(100)
' Set q = BigInteger.DivRem(b, BInt(40), r)
' Debug.Print q.ToString '' 2
' Debug.Print r.ToString '' 20
' </pre>
Public Function DivRem(ByVal Dividend As BigInteger, ByVal Divisor As BigInteger, ByRef Remainder As BigInteger) As BigInteger
    Set Dividend = ValueOrZero(Dividend)
    Set DivRem = Dividend.DivRem(Divisor, Remainder)
End Function

''
' Returns the remainder of division between two <a href="BigInteger.html">BigInteger</a> values.
'
' @param dividend The value to be divided.
' @param divisor The value to divide by.
' @return Returns the remainder portion of the division operation.
' @remarks Values of <b>Nothing</b> will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BInt(100)
' Set b = BigInteger.Remainder(b, BInt(40))
' Debug.Print b.ToString '' 20
' </pre>
Public Function Modulus(ByVal Dividend As BigInteger, ByVal Divisor As BigInteger) As BigInteger
    Set Dividend = ValueOrZero(Dividend)
    Set Modulus = Dividend.Modulus(Divisor)
End Function

''
' Returns the greatest common divisor between the two <a href="BigInteger.html">BigInteger</a> values.
'
' @param Left One of the values to find the GCD of.
' @param Right One of the values to find the GCD of.
' @return Returns the GCD for two given values.
' @remarks Values of <b>Nothing</b> will be treated as zero.
Public Function GreatestCommonDivisor(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set Right = ValueOrZero(Right)
    
    If Left.IsZero And Right.IsZero Then
        Set GreatestCommonDivisor = BigInteger.Zero
    Else
        Set Left = Left.Absolute
        Set Right = Right.Absolute
    
        Dim Result As BigInteger
        Do While Not Right.IsZero
            Set Result = Left.Modulus(Right)
            Set Left = Right
            Set Right = Result
        Loop
    
        Set GreatestCommonDivisor = Left
    End If
End Function

''
' Tests a specific bit in the <a href="BigInteger.html">BigInteger</a> number to determine if it is set (1) or not (0).
'
' @param Value The <a href="BigInteger.html">BigInteger</a> number to test the bit of.
' @param bit The bit to be tested.
' @return Returns True if the bit is set (1), false if it is not (0).
' @remarks The max value for bit is 2147483647. If the bit value is larger than
' the <a href="BigInteger.html">BigInteger</a> instance, then the sign of the number is used. This will represent
' a negative number with infinite sign extension bits.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function TestBit(ByVal Value As BigInteger, ByVal Bit As Long) As Boolean
    Set Value = ValueOrZero(Value)
    TestBit = Value.TestBit(Bit)
End Function

''
' Sets the specified bit of the <a href="BigInteger.html">BigInteger</a> instance to 1.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> value to set the bit on.
' @param bit The bit to be set to 1.
' @return A new <a href="BigInteger.html">BigInteger</a> value with the bit set to 1.
' @remarks If the bit specified is larger than the current instance
' of the <a href="BigInteger.html">BigInteger</a> value, then the new value will be extended to include the bit being set.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function SetBit(ByVal Value As BigInteger, ByVal Bit As Long) As BigInteger
    Set Value = ValueOrZero(Value)
    SetBit = Value.SetBit(Bit)
End Function

''
' Clears the specified bit of the <a href="BigInteger.html">BigInteger</a> instance to 0.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> value to clear the bit on.
' @param bit The bit to be set to 0.
' @return A new <a href="BigInteger.html">BigInteger</a> value with the bit set to 0.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function ClearBit(ByVal Value As BigInteger, ByVal Bit As Long) As BigInteger
    Set Value = ValueOrZero(Value)
    Set ClearBit = Value.ClearBit(Bit)
End Function

''
' Flips the specified bit, toggling between True and False.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> value to flip the bit on.
' @param bit The bit to toggle.
' @return A new <a href="BigInteger.html">BigInteger</a> with the specified bit flipped.
' @remarks If the bit specified is larger than the current instance
' of the <a href="BigInteger.html">BigInteger</a> value, then the new value will be extended to
' include the bit being flipped.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function FlipBit(ByVal Value As BigInteger, ByVal Bit As Long) As BigInteger
    Set Value = ValueOrZero(Value)
    Set FlipBit = Value.FlipBit(Bit)
End Function

''
' Peforms a bitwise AND operation, returning the result.
'
' @param Left The left hand value of the AND operation.
' @param Right The right hand value of the AND operation.
' @return A new <a href="BigInteger.html">BigInteger</a> value containing the performed bitwise AND result.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function BitwiseAnd(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set BitwiseAnd = Left.BitwiseAnd(Right)
End Function

''
' Peforms a bitwise OR operation, returning the result.
'
' @param Left The left hand value of the OR operation.
' @param Right The right hand value of the OR operation.
' @return A new <a href="BigInteger.html">BigInteger</a> value containing the performed bitwise OR result.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function BitwiseOr(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set BitwiseOr = Left.BitwiseOr(Right)
End Function

''
' Peforms a bitwise XOR operation, returning the result.
'
' @param Left The left hand value of the XOR operation.
' @param Right The right hand value of the XOR operation.
' @return A new <a href="BigInteger.html">BigInteger</a> value containing the performed bitwise XOR result.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function BitwiseXor(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set BitwiseXor = Left.BitwiseXor(Right)
End Function

''
' Peforms a bitwise NOT of this <a href="BigInteger.html">BigInteger</a> instance.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> to perform the NOT operation on.
' @return A new <a href="BigInteger.html">BigInteger</a> value containing the performed bitwise NOT result.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function BitwiseNot(ByVal Value As BigInteger) As BigInteger
    Set Value = ValueOrZero(Value)
    Set BitwiseNot = Value.BitwiseNot
End Function

''
' Peforms a bitwise AND operation after NOTting the second operand, returning the result.
'
' @param Left The left hand value of the operation.
' @param Right The right hand value to be NOTted for the operation.
' @return A new <a href="BigInteger.html">BigInteger</a> value containing the performed bitwise result.
' <p>Values of Nothing will be treated as zero.</p>
Public Function BitwiseAndNot(ByVal Left As BigInteger, ByVal Right As BigInteger) As BigInteger
    Set Left = ValueOrZero(Left)
    Set BitwiseAndNot = Left.BitwiseAndNot(Right)
End Function

''
' Raise a base value to the specified power.
'
' @param Value The value to be raised to a specified power.
' @param Exponent The power to raise the value to.
' @return Returns the baseValue^exponent result.
' @remarks Values of <b>Nothing</b> are treated as zero.
Public Function Pow(ByVal Value As BigInteger, ByVal Exponent As Long) As BigInteger
    If Exponent < 0 Then _
        Error.ArgumentOutOfRange "Exponent", ArgumentOutOfRange_NeedNonNegNum
    
    Set Value = ValueOrZero(Value)
    
    Select Case Exponent
        Case 0
            Set Pow = mOne
        Case 1
            Set Pow = Value
        Case Else
            Set Pow = CalculatePow(Value, Exponent)
    End Select
End Function

Private Function CalculatePow(ByVal Value As BigInteger, ByVal Exponent As Long) As BigInteger
    Dim y       As BigInteger
    Dim z       As BigInteger
    Dim IsOdd   As Boolean
    
    Set y = mOne
    Set z = Value
    
    Do
        IsOdd = Exponent And 1
        Exponent = Exponent \ vbShift1Bit
    
        If IsOdd Then
            Set y = Multiply(y, z)
            
            If Exponent = 0 Then
                Set CalculatePow = y
                Exit Function
            End If
        End If
        
        Set z = Multiply(z, z)
    Loop
End Function

''
' Performs a modulo on a <a href="BigInteger.html">BigInteger</a> value raised to a power. b = (x^y) mod z.
'
' @param baseValue The value to be raised to the power of the exponent.
' @param exponent The power to raise the baseValue to.
' @param m The modulus value to perform on the raised baseValue.
' @return Returns the modulus of the raised baseValue.
' <p>Values of <p>Nothing</p> are treated as zero.</p>
Public Function ModPow(ByVal BaseValue As BigInteger, ByVal Exponent As BigInteger, ByVal m As BigInteger) As BigInteger
    Set BaseValue = Pow(BaseValue, Exponent)
    DivRem BaseValue, m, ModPow
End Function

''
' Returns the mininum of the two <a href="BigInteger.html">BigInteger</a> values.
'
' @param x One of the two values to be tested.
' @param y One of the two values to be tested.
' @return If x is less or equal to y, then x is returned, otherwise y is returned.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function Min(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    If Compare(x, y) <= 0 Then
        Set Min = x
    Else
        Set Min = y
    End If
End Function

''
' Returns the maximum of the two <a href="BigInteger.html">BigInteger</a> values.
'
' @param x One of the two values to be tested.
' @param y One of the two values to be tested.
' @return If x is equal or greater than y, then x is returned, otherwise y is returned.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function Max(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    If Compare(x, y) >= 0 Then
        Set Max = x
    Else
        Set Max = y
    End If
End Function

''
' Randomizes the Rnd function.
'
' @param Value The number used to randomize the Rnd method.
' @remarks This is provided because the VB Rnd function runs
' independently for each component.
Public Sub Randomize(Optional ByRef Value As Variant)
    VBA.Randomize Value
End Sub

''
' Sets the seeding of the Rnd function.
'
' @param seed The number to be used as the seed.
' @remarks This is provided because the VB Rnd function runs
' independently for each component.
Public Sub SeedRnd(Optional ByRef seed As Variant)
    VBA.Rnd seed
End Sub

''
' Generates a random <a href="BigInteger.html">BigInteger</a> number of the specified size.
'
' @param size The size of the number in bytes.
' @param IsNegative Indicates if a negative random number should be generated.
' @return A new <a href="BigInteger.html">BigInteger</a> of a random value.
Public Function Rnd(ByVal Size As Long, Optional ByVal IsNegative As Boolean = False) As BigInteger
    If Size < 0 Then _
        Error.ArgumentOutOfRange "Size", ArgumentOutOfRange_NeedNonNegNum
    
    If Size = 0 Then
        Set Rnd = mZero
    Else
        Set Rnd = New BigInteger
        Rnd.InitRnd Size, IsNegative
    End If
End Function

''
' Performs a left shift operation on the <a href="BigInteger.html">BigInteger</a> instance.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> to shift left.
' @param BitCount The number of bits the value is to be shifted left.
' @return Returns a new <a href="BigInteger.html">BigInteger</a> instance containing the shifted value.
' @remarks If BitCount is negative, then a right shift is performed.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function ShiftLeft(ByVal Value As BigInteger, ByVal BitCount As Long) As BigInteger
    Set Value = ValueOrZero(Value)
    Set ShiftLeft = Value.ShiftLeft(BitCount)
End Function

''
' Performs a bit shifting operation to the the right.
'
' @param Value The <a href="BigInteger.html">BigInteger</a> to shift right.
' @param BitCount The number of bits the <a href="BigInteger.html">BigInteger</a> value is to be shifted.
' @return Returns a new <a href="BigInteger.html">BigInteger</a> instance containing the bit shifted value.
' @remarks If BitCount is negative, then a left shift is performed.
' <p>Values of <b>Nothing</b> will be treated as zero.</p>
Public Function ShiftRight(ByVal Value As BigInteger, ByVal BitCount As Long) As BigInteger
    Set Value = ValueOrZero(Value)
    Set ShiftRight = Value.ShiftRight(BitCount)
End Function

''
' Returns the integer square root of a <a href="BigInteger.html">BigInteger</a>.
'
' @param Value The value to calculate the integer square root for.
' @Return The integer square root of the value.
Public Function SquareRoot(ByVal Value As BigInteger) As BigInteger
    Set Value = ValueOrZero(Value)
    
    If Value.Sign = 0 Then
        Set SquareRoot = mZero
        Exit Function
    End If
    
    If Value.Sign = -1 Then
        Set Value = Value.Negate
    End If
    
    Dim Estimate As BigInteger
    Dim n As BigInteger
    
    Set Estimate = Value.ShiftRight(1).Add(mOne)
    Set n = Value.Divide(Estimate).Add(Estimate).ShiftRight(1)
    
    Do While Compare(n, Estimate) < 0
        Set Estimate = n
        Set n = Value.Divide(Estimate).Add(Estimate).ShiftRight(1)
    Loop
    
    Set SquareRoot = Estimate
End Function

''
' Calculates the factorial of a <a href="BigInteger.html">BigInteger</a> value.
'
' @param Value The value to calculate the factorial of.
' @return A <a href="BigInteger.html">BigInteger</a> value of the factorial.
Public Function Factorial(ByVal Value As BigInteger) As BigInteger
    Dim Result As BigInteger
    
    Set Value = ValueOrZero(Value)
    
    If Value.Sign = Negative Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_NeedNonNegNum
    
    Set Result = mOne
    
    Do While Not Value.IsZero
        Set Result = Multiply(Result, Value)
        Set Value = Subtract(Value, mOne)
    Loop
    
    Set Factorial = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mZero = Cor.NewBigInteger(0)
    Set mOne = Cor.NewBigInteger(1)
    Set mMinusOne = Cor.NewBigInteger(-1)
    Set mTen = Cor.NewBigInteger(10)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValueOrZero(ByVal Value As BigInteger) As BigInteger
    If Value Is Nothing Then
        Set ValueOrZero = mZero
    Else
        Set ValueOrZero = Value
    End If
End Function

Private Function TryParseDecimal(ByRef s As String, ByVal Style As NumberStyles, ByVal Info As NumberFormatInfo, ByRef n As BigNumber) As Boolean
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
    
    NegativeSign = AscW(Info.NegativeSign)
    PositiveSign = AscW(Info.PositiveSign)
    CurrencySymbol = AscW(Info.CurrencySymbol)
    DecimalPoint = AscW(Info.NumberDecimalSeparator)
    CurrencyPoint = AscW(Info.CurrencyDecimalSeparator)
    Thousands = AscW(Info.NumberGroupSeparator)
    CurrencyThousands = AscW(Info.CurrencyGroupSeparator)
    Chars = AllocChars(s)
    MaxIndex = UBound(Chars)
    n.Sign = 1
    ReDim n.Digits(0 To Len(s) \ 4)
    
    Do While Index <= MaxIndex
        Ch = Chars(Index)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                If (FoundCurrency = FoundAfter) Or (FoundSign = FoundAfter) Or InWhitespace Then
                    GoTo Finally
                End If
                
                If Not FoundDecimal Then
                    SingleInPlaceMultiply n, 10
                    SingleInPlaceAdd n, Ch - vbZeroChar
                Else
                    If Style And NumberStyles.AllowExponent Then
                        ' we are allowing a possible exponent value to be parsed,
                        ' so if we have encountered a decimal point, we need to
                        ' continue parsing digits and also count how many
                        ' decimal places we have gone.
                        SingleInPlaceMultiply n, 10
                        SingleInPlaceAdd n, Ch - vbZeroChar
                        
                        DecimalPlaces = DecimalPlaces + 1
                    Else
                        If Ch <> vbZeroChar Then
                            GoTo Finally
                        End If
                    End If
                End If
                
                FoundDigit = True
            Case NegativeSign, PositiveSign
                If (FoundSign <> NotFound) Or NeedsClosingParen Then
                    GoTo Finally
                End If
                
                If Not FoundDigit Then
                    If (Style And NumberStyles.AllowLeadingSign) = 0 Then
                        GoTo Finally
                    End If
                    
                    FoundSign = FoundBefore
                Else
                    If (Style And NumberStyles.AllowTrailingSign) = 0 Then
                        GoTo Finally
                    End If
                    
                    FoundSign = FoundAfter
                End If
                
                If Ch = NegativeSign Then
                    n.Sign = -1
                End If
            Case vbTabChar, vbLfChar, vbVtChar, vbFfChar, vbCrChar, vbSpaceChar
                If Not FoundDigit Then
                    If ((Style And NumberStyles.AllowLeadingWhite) = 0) Or NeedsClosingParen Then
                        GoTo Finally
                    End If
                
                    InWhitespace = FoundSign
                Else
                    InWhitespace = True
                End If
            Case vbOpenParenChar
                If ((Style And NumberStyles.AllowParentheses) = 0) Or (FoundSign <> NotFound) Then
                    GoTo Finally
                End If
                
                NeedsClosingParen = True
            Case vbCloseParenChar
                If (Style And NumberStyles.AllowParentheses) = 0 Or (Not NeedsClosingParen) Then
                    GoTo Finally
                End If
                
                n.Sign = -1
                NeedsClosingParen = False
                FoundClosingParen = True
            Case CurrencySymbol
                If ((Style And NumberStyles.AllowCurrencySymbol) = 0) Or (FoundCurrency <> NotFound) Then
                    GoTo Finally
                End If
                
                FoundCurrency = IIfLong(n.Precision = 0, FoundBefore, FoundAfter)
            Case DecimalPoint, CurrencyPoint
                If ((Style And NumberStyles.AllowDecimalPoint) = 0) Or FoundDecimal Or FoundClosingParen Then
                    GoTo Finally
                End If
                
                If FoundCurrency Then
                    If Ch = DecimalPoint Then
                        GoTo Finally
                    End If
                End If
                
                FoundDecimal = True
            Case Thousands, CurrencyThousands
                If ((Style And NumberStyles.AllowThousands) = 0) Or (Not FoundDigit) Then
                    GoTo Finally
                End If
                
                If FoundCurrency Then
                    If Ch = Thousands Then
                        GoTo Finally
                    End If
                End If
            Case vbLowerEChar, vbUpperEChar
                If ((Style And NumberStyles.AllowExponent) = 0) Or FoundExponent Then
                    GoTo Finally
                End If
            
                If Not TryParseExponent(Chars, Index, Exponent) Then
                    GoTo Finally
                End If
                
                ' we understand that parsing the exponent will advance the index passed the end of the digits
                ' and that we will be advancing the index as well, in effect advancing it twice, so we need
                ' to counter that last advancement.
                Index = Index - 1
                FoundExponent = True
            Case Else
                GoTo Finally
        End Select
        
        Index = Index + 1
    Loop
    
    If InWhitespace Then
        If (Style And NumberStyles.AllowTrailingWhite) = 0 Then
            GoTo Finally
        End If
    End If
    
    If NeedsClosingParen Then
        GoTo Finally
    End If
    
    If FoundExponent Then
        If Exponent < DecimalPlaces Then
            GoTo Finally
        End If
    
        n.Precision = n.Precision + Exponent \ 4
        ReDim Preserve n.Digits(0 To n.Precision)
        
        Do While Exponent - DecimalPlaces > 0
            SingleInPlaceMultiply n, 10
            Exponent = Exponent - 1
        Loop
    End If
    
    Select Case n.Precision
        Case 0
            Erase n.Digits
            n.Sign = 0
        Case Is < UBound(n.Digits) + 1
            ReDim Preserve n.Digits(0 To n.Precision - 1)
    End Select
    
    If n.Sign = -1 Then
        ApplyTwosComplement n.Digits
    End If
    
    TryParseDecimal = True
    
Finally:
    FreeChars Chars
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

Private Function TryParseHex(ByRef s As String, ByVal HasPrefix As Boolean, ByVal Style As NumberStyles, ByRef n As BigNumber) As Boolean
    Const vbLowerHexOffset As Long = 87
    Const vbUpperHexOffset As Long = 55
    Dim Chars()     As Integer
    Dim MaxSize     As Long
    Dim MinIndex    As Long
    Dim UpperIndex  As Long
    Dim FoundDigits As Boolean
    Dim Length      As Long
    
    MinIndex = IIfLong(HasPrefix, 2, 0)
    Length = Len(s) - MinIndex
    Chars = AllocChars(s)
    UpperIndex = UBound(Chars)
    MaxSize = (Length) \ 4 + IIfLong(Length Mod 4, 1, 0)
    ReDim n.Digits(0 To MaxSize - 1)
    
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Ch As Long
    
    j = 1
    For i = UpperIndex To MinIndex Step -1
        Select Case Chars(i)
            Case vbZeroChar
                Ch = 0
                FoundDigits = True
            Case vbOneChar To vbNineChar
                Ch = Chars(i) - vbZeroChar
                n.Sign = 1
                FoundDigits = True
            Case vbLowerAChar To vbLowerFChar
                Ch = Chars(i) - vbLowerHexOffset
                n.Sign = 1
                FoundDigits = True
            Case vbUpperAChar To vbUpperFChar
                Ch = Chars(i) - vbUpperHexOffset
                n.Sign = 1
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
        
        k = k + Ch * j
        j = j * vbShift4Bits
        
        If j = &H10000 Then
            #If Release Then
                n.Digits(c) = k
            #Else
                n.Digits(c) = AsWord(k)
            #End If
            
            k = 0
            j = 1
            c = c + 1
        End If
        
Continue:
    Next i
    
    If k Then
        #If Release Then
            n.Digits(c) = k
        #Else
            n.Digits(c) = AsWord(k)
        #End If

        c = c + 1
    End If
    
    n.Precision = c
    
    If n.Sign Then
        Ch = n.Digits(n.Precision - 1)
            
        If (Ch And &H8000) = &H8000 Then
            n.Sign = -1
        ElseIf (Ch And &HF800) = &H800 Then
            n.Sign = -1
            n.Digits(n.Precision - 1) = Ch Or &HF000
        ElseIf (Ch And &HFF80) = &H80 Then
            n.Sign = -1
        ElseIf (Ch And &HFFF8) = &H8 Then
            n.Sign = -1
            n.Digits(n.Precision - 1) = Ch Or &HFFF0
        End If
    End If
    
    TryParseHex = True
    
Finally:
    FreeChars Chars
End Function
