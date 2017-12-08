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
' Provides methods used to manipulate BigInteger objects.
'
' @remarks The methods provided by this class are accessed through the
' convention of <i><b>BigInteger.MethodName</b></i>.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Rnd(128)
' Set b = BigInteger.Pow(b, BInt(3))
' </pre>
' <p>Most of these methods exist in the <b>BitInteger</b> class and can be used
' directly with <b>BigInteger</b> objects. The primary advantage of using the
' methods defined in this class is that a value of Nothing can be passed in and
' will be treated as a value of 0.
'
Option Explicit

Private mZero       As BigInteger
Private mOne        As BigInteger
Private mMinusOne   As BigInteger
Private mTen        As BigInteger


''
' Returns a shared BigInteger instance with the value of 0.
'
' @return Returns a shared instance of the value of 0.
' @remarks This is to reduce the creation of multiple instances
' of values that are commonly used.
'
Public Property Get Zero() As BigInteger
    Set Zero = mZero
End Property

''
' Returns a shared BigInteger instance with the value of 1.
'
' @return Returns a shared instance of the value of 1.
' @remarks This is to reduce the creation of multiple instances
' of values that are commonly used.
'
Public Property Get One() As BigInteger
    Set One = mOne
End Property

''
' Returns a shared BigInteger instance with the value of -1.
'
' @return Returns a shared instance of the value of -1.
' @remarks This is to reduce the creation of multiple instances
' of values that are commonly used.
'
Public Property Get MinusOne() As BigInteger
    Set MinusOne = mMinusOne
End Property

''
' Returns a shared BigInteger instance with the value of 10.
'
' @return Returns a shared instance of the value of 10.
' @remarks This is to reduce the creation of multiple instances
' of values that are commonly used.
'
Public Property Get Ten() As BigInteger
    Set Ten = mTen
End Property

''
' Compares two BigInteger instance to each other returning their relation.
'
' @param x The left handed value to be compared to.
' @param y The right handed value to be compared to.
' @return Returns the relation between to the two numbers.<br>
' @remarks If Nothing is passed in, it is treated as zero.<br>
' If x < y then -1 is returned<br>
' If x = y then 0 is returned<br>
' If x > y then 1 is returned.<br>
'
Public Function Compare(ByVal x As BigInteger, ByVal y As BigInteger) As Long
    ZeroIfNothing x
    ZeroIfNothing y
    
    Dim xNumber As Number
    Dim yNumber As Number
    Dim Result  As Long
    
    x.AttachNumber xNumber
    y.AttachNumber yNumber
    Result = xNumber.Sign - yNumber.Sign
    
    If Result = 0 Then
        Result = xNumber.Precision - yNumber.Precision
        
        If Result = 0 Then
            Dim i As Long
            For i = xNumber.Precision - 1 To 0 Step -1
                Result = (xNumber.Digits(i) And &HFFFF&) - (yNumber.Digits(i) And &HFFFF&)
                If Result <> 0 Then Exit For
            Next i
        Else
            Result = xNumber.Sign * Result
        End If
    End If
    
    x.DetachNumber xNumber
    y.DetachNumber yNumber
    
    Compare = Sgn(Result)
End Function

''
' Returns the equality of two BigInteger values.
'
' @param x The left handed value to test for equality.
' @param y The right handed value to test for equality.
' @return Returns True if the values are equal, False otherwise.
' @remarks If Nothing is passed in, it is treated as zero.
'
Public Function Equals(ByVal x As BigInteger, ByVal y As BigInteger) As Boolean
    Equals = (Compare(x, y) = 0)
End Function

''
' Attemps to parse a string of characters of the specified base type.
'
' @param Value The string of characters to attempt to be parsed.
' @param OutResult The returning BigInteger object if the string was successfully parsed.
' @return Returns True if the string was parsed, False otherwise.
' @remarks 3 base types are supported for parsing (Decimal, Hex, Binary.) The string
' must include a base type specifier for Hex and Binary. Decimal is the default.
' <p>
' Hex Specifiers: 0x, 0X, &h, &H - (0xFF, 0XFF, &hFF, &HFF) <br>
' Binary Specifiers: 0b, 0B - (0b00001111, 0B000011111) <br>
' </p>
' Anything else will default to Decimal.<br>
' "-" is supported for all types of parsing.
' <pre>
' Dim b As BigInteger
' Debug.Print BigInteger.TryParse("-&H7FFF", b) '' True
' Debug.Print b.ToString("X") '' 8001
' </pre>
' <p>Parsing is not case sensitive.</p>
Public Function TryParse(ByRef Value As String, ByRef OutResult As BigInteger) As Boolean
    TryParse = TryParseEx(Value, NumberStyles.Integer, Nothing, OutResult)
End Function

''
' Tries to convert the string representation of a number in a specified style and culture-specific format to its BigInteger equivalent, and returns a value that indicates whether the conversion succeeded.
'
' @param Value The string representation of a number. The string is interpreted using the style specified by style.
' @param Style A bitwise combination of enumeration values that indicates the style elements that can be present in value. A typical value to specify is NumberStyles.Integer.
' @param Provider An object that supplies culture-specific formatting information about value. If Nothing is passed in the current CultureInfo will be used.
' @param OutResult When this method returns, contains the BigInteger equivalent to the number that is contained in value, or BigInteger.Zero if the conversion failed. The conversion fails if thevalue parameter is null or is not in a format that is compliant withstyle. This parameter is passed uninitialized.
Public Function TryParseEx(ByRef Value As String, ByVal Style As NumberStyles, ByVal Provider As IFormatProvider, ByRef OutResult As BigInteger) As Boolean
    Dim Trimmed As String
    Dim Result  As Number
    Dim Info    As NumberFormatInfo
    
    Set Info = NumberFormatInfo.GetInstance(Provider)
    
    If TryParseDecimal(Value, Style, Info, Result) Then
        Set OutResult = New BigInteger
        OutResult.InternalInitNumber Result
        TryParseEx = True
    Else
        Set OutResult = Nothing
    End If
End Function

''
' Attemps to parse a string of characters of the specified base type.
'
' @param Value The string of characters to attempt to be parsed.
' @return Returns an instance of BigInteger containing the parsed value.
' @remarks If the string could not be parsed, an exception is thrown.
' <p>3 base types are supported for parsing (Decimal, Hex, Binary.) The string
' must include a base type specifier for Hex and Binary. Decimal is the default.</p>
' <p>
' Hex Specifiers: 0x, 0X, &h, &H - (0xFF, 0XFF, &hFF, &HFF) <br>
' Binary Specifiers: 0b, 0B - (0b00001111, 0B000011111) <br>
' </p>
' Anything else will default to Decimal.<br>
' "-" is supported for all types of parsing.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Parse("-&H7FFF")
' Debug.Print b.ToString("X") '' 8001
' </pre>
' <p>Parsing is not case sensitive.</p>
'
Public Function Parse(ByRef Value As String, Optional ByVal Style As NumberStyles = NumberStyles.Integer, Optional ByVal Provider As IFormatProvider) As BigInteger
    If Not TryParseEx(Value, Style, Provider, Parse) Then
        Error.Format Format_ParseBigInteger
    End If
End Function

''
' Returns a human readable string of characters representing the number.
'
' @param x The BigInteger to create a string from.
' @param format The specific numerical format wish to have the number displayed as.
' @return The numeric value of the number in the specified format.
' @remarks "X" and "x" for hex in either upper or lowercase. "B" or "b" for a binary
' representation of 1's and 0's. Everything else defaults to a decimal display.
' <p>The specifier can also contain a number of mininum digits to be displayed. If
' the value isn't long enough to reach the mininum digits, then leading 0's, 1's, or
' FF's will be added to the beginning of the output.</p>
' <pre>
' Dim b As BigInteger
' Set b = BInt(-200)
' Debug.Print b.ToString("d8") '' -00000200
' Debug.Print b.ToString("d")  '' -200
' Debug.Print b.ToString("x8") '' ffffff38
' Debug.Print b.ToString("x")  '' ff38
' Debug.Print b.ToString("b8") '' 1111111100111000
' Debug.Print b.ToString("b")  '' 1111111100111000
' </pre>
' <p>If no format is specified, the default of "d" is used.</p>
'
Public Function ToString(ByVal x As BigInteger, Optional ByVal Format As String) As String
    ZeroIfNothing x
    ToString = x.ToString(Format)
End Function

''
' Returns the opposite signed value of this BigInteger instance.
'
' @param x The value to be negated.
' @return Returns a BigInteger instance containing the opposite signed value.
' @remarks Negative values are represented as a Twos Complement number, unlike
' other systems used to represent a negative value. Any negation requires
' the Twos Complement algorithm to be applied to the number. This keeps in step
' with how users of VB already deal with negative numbers, helping to ease
' the transition to using BigInteger values.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function Negate(ByVal x As BigInteger) As BigInteger
    ZeroIfNothing x
    Set Negate = x.Negate
End Function

''
' Returns the absolute (positive or zero) value of this instance.
'
' @param x The BigInteger value to return the absolute value of.
' @return Returns a positive (or zero) value of the current BigInteger instance.
' @remarks If the value is already positive or zero, then the same object instance
' is returned. If the value is negative, then the negated value is returned.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function Absolute(ByVal x As BigInteger) As BigInteger
    ZeroIfNothing x
    Set Absolute = x.Absolute
End Function

''
' Adds two BigInteger value instances together, returning the sum.
'
' @param x The augend to be added to.
' @param y The addend to add to the augend.
' @return The sum of the two values.
' @remarks Values of Nothing will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Add(BInt(100), BInt(150))
' Debug.Print b.ToString '' 250
' </pre>
'
Public Function Add(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    ZeroIfNothing x
    ZeroIfNothing y
    
    If x.Sign = Sign.Zero Then
        Set Add = y
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set Add = x
        Exit Function
    End If
    
    Dim u       As Number
    Dim v       As Number
    Dim sum()   As Integer
    
    x.AttachNumber u
    y.AttachNumber v
    sum = GradeSchoolAdd(u, v)
    x.DetachNumber u
    y.DetachNumber v
    
    Set Add = New BigInteger
    Add.InternalInit sum
End Function

''
' Subtracts one BigInteger value from another, returning the difference.
'
' @param x The minuend value to be subracted from.
' @param y The subtrahend value to subtract from the minuend value.
' @return The difference of the two values.
' @remarks Values of Nothing will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Subtract(BInt(100), BInt(150))
' Debug.Print b.ToString '' -50
' </pre>
'
Public Function Subtract(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    ZeroIfNothing x
    ZeroIfNothing y
    
    If x.Sign = Sign.Zero Then
        Set Subtract = y.Negate
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set Subtract = x
        Exit Function
    End If
    
    Dim u               As Number
    Dim v               As Number
    Dim Difference()    As Integer
    
    x.AttachNumber u
    y.AttachNumber v
    Difference = GradeSchoolSubtract(u, v)
    x.DetachNumber u
    y.DetachNumber v
    
    Set Subtract = New BigInteger
    Subtract.InternalInit Difference
End Function

''
' Multiplies BigInteger values, returning their product.
'
' @param x The multiplicand to be multiplied.
' @param y The multiplier to multiply by.
' @return The product of the two values.
' @remarks Values of Nothing will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Multiply(BInt(100), BInt(150))
' Debug.Print b.ToString '' 15000
' </pre>
'
Public Function Multiply(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    ZeroIfNothing x
    ZeroIfNothing y
    
    If (x.Sign = Sign.Zero) Or (y.Sign = Sign.Zero) Then
        Set Multiply = mZero
        Exit Function
    ElseIf x.IsOne Then
        Set Multiply = y
        Exit Function
    ElseIf y.IsOne Then
        Set Multiply = x
        Exit Function
    ElseIf x.IsMinusOne Then
        Set Multiply = y.Negate
        Exit Function
    ElseIf y.IsMinusOne Then
        Set Multiply = x.Negate
        Exit Function
    End If
    
    Dim u As Number
    Dim v As Number
    
    If x.Sign = Sign.Negative Then
        x.CopyNumber u
        Mathematics.Negate u
    Else
        x.AttachNumber u
    End If

    If y.Sign = Sign.Negative Then
        y.CopyNumber v
        Mathematics.Negate v
    Else
        y.AttachNumber v
    End If
    
    Dim Product() As Integer
    Product = GradeSchoolMultiply(u, v)
    
    If x.Sign = Sign.Positive Then
        x.DetachNumber u
    End If
    If y.Sign = Sign.Positive Then
        y.DetachNumber v
    End If
    If x.Sign <> y.Sign Then
        ApplyTwosComplement Product
    End If
        
    Set Multiply = New BigInteger
    Multiply.InternalInit Product
End Function


''
' Divides one BigInteger value by another, returning the quotient.
'
' @param dividend The BigInteger value to be divided.
' @param divisor The BigInteger value to divide by.
' @return The quotient of the division.
' @remarks Values of Nothing will be treated as zero.
' <p>Integer division is applied, there is no rounding values. This method discards the remainder.</p>
' <pre>
' Dim b As BigInteger
' Set b = BigInteger.Divide(BInt(100), BInt(40))
' Debug.Print b.ToString '' 2
' </pre>
'
Public Function Divide(ByVal Dividend As BigInteger, ByVal Divisor As BigInteger) As BigInteger
    Set Divide = InternalDivide(Dividend, Divisor, Nothing, False)
End Function

''
' Divides one BigInteger value by another, returning both the quotient and remainder.
'
' @param dividend The BigInteger value to be divided.
' @param divisor The BigInteger value to divide by.
' @param remainder A variable to place the remainder in to.
' @return The quotient of the division.
' @remarks Values of Nothing will be treated as zero.
' <pre>
' Dim b As BigInteger
' Dim r As BigInteger
' Dim q As BigInteger
' Set b = BInt(100)
' Set q = BigInteger.DivRem(b, BInt(40), r)
' Debug.Print q.ToString '' 2
' Debug.Print r.ToString '' 20
' </pre>
Public Function DivRem(ByVal Dividend As BigInteger, ByVal Divisor As BigInteger, ByRef remainder As BigInteger) As BigInteger
    Set DivRem = InternalDivide(Dividend, Divisor, remainder, True)
End Function

''
' Returns the modulus (remainder) of division between two BigInteger values.
'
' @param dividend The value to be divided.
' @param divisor The value to divide by.
' @return Returns the remainder portion of the division operation.
' @remarks Values of Nothing will be treated as zero.
' <pre>
' Dim b As BigInteger
' Set b = BInt(100)
' Set b = BigInteger.Modulus(b, BInt(40))
' Debug.Print b.ToString '' 20
' </pre>
Public Function Modulus(ByVal Dividend As BigInteger, ByVal Divisor As BigInteger) As BigInteger
    InternalDivide Dividend, Divisor, Modulus, True
End Function

''
' Returns the greatest common divisor between the two BigInteger values.
'
' @param x One of the values to find the GCD of.
' @param y One of the values to find the GCD of.
' @return Returns the GCD for two given values.
' @remarks Values of Nothing will be treated as zero.
'
Public Function GreatestCommonDivisor(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    ZeroIfNothing x
    ZeroIfNothing y
    
    If x.Sign = Sign.Zero And y.Sign = Sign.Zero Then
        Set GreatestCommonDivisor = mZero
        Exit Function
    End If
    
    Set x = x.Absolute
    Set y = y.Absolute
    
    Dim R As BigInteger
    Do While y.Sign <> Sign.Zero
        Set R = Modulus(x, y)
        Set x = y
        Set y = R
    Loop

    Set GreatestCommonDivisor = x
End Function

''
' Tests a specific bit in the BigInteger number to determine if it is set (1) or not (0).
'
' @param x The BigInteger number to test the bit of.
' @param bit The bit to be tested.
' @return Returns True if the bit is set (1), false if it is not (0).
' @remarks The max value for bit is 2147483647. If the bit value is larger than
' the BigInteger instance, then the sign of the number is used. This will represent
' a negative number with infinite sign extension bits.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function TestBit(ByVal x As BigInteger, ByVal Bit As Long) As Boolean
    ZeroIfNothing x
    TestBit = x.TestBit(Bit)
End Function

''
' Sets the specified bit of the BigInteger instance to 1.
'
' @param x The BigInteger value to set the bit on.
' @param bit The bit to be set to 1.
' @return A new BigInteger value with the bit set to 1.
' @remarks If the bit specified is larger than the current instance
' of the BigInteger value, then the new value will be extended to
' include the bit being set.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function SetBit(ByVal x As BigInteger, ByVal Bit As Long) As BigInteger
    ZeroIfNothing x
    SetBit = x.SetBit(Bit)
End Function

''
' Clears the specified bit of the BigInteger instance to 0.
'
' @param x The BigInteger value to clear the bit on.
' @param bit The bit to be set to 0.
' @return A new BigInteger value with the bit set to 0.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function ClearBit(ByVal x As BigInteger, ByVal Bit As Long) As BigInteger
    ZeroIfNothing x
    Set ClearBit = x.ClearBit(Bit)
End Function

''
' Flips the specified bit, toggling between True and False.
'
' @param x The BigInteger value to flip the bit on.
' @param bit The bit to toggle.
' @return A new BigInteger with the specified bit flipped.
' @remarks If the bit specified is larger than the current instance
' of the BigInteger value, then the new value will be extended to
' include the bit being flipped.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function FlipBit(ByVal x As BigInteger, ByVal Bit As Long) As BigInteger
    ZeroIfNothing x
    Set FlipBit = x.FlipBit(Bit)
End Function

''
' Peforms a bitwise AND operation, returning the result.
'
' @param x The left hand value of the AND operation.
' @param y The right hand value of the AND operation.
' @return A new BigInteger value containing the performed bitwise AND result.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function AndBits(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    ZeroIfNothing x
    ZeroIfNothing y
    
    If x.Sign = Sign.Zero Or y.Sign = Sign.Zero Then
        Set AndBits = mZero
        Exit Function
    End If
    
    Dim ExtDigit    As Integer
    Dim Number()    As Integer
    Dim LongNumber  As Number
    Dim ShortNumber As Number
    
    If x.Precision >= y.Precision Then
        x.AttachNumber LongNumber
        y.AttachNumber ShortNumber
    Else
        x.AttachNumber ShortNumber
        y.AttachNumber LongNumber
    End If
    
    If ShortNumber.Sign = Sign.Negative Then
        ExtDigit = &HFFFF
    End If
    
    ReDim Number(0 To LongNumber.Precision)
    
    Dim i As Long
    For i = 0 To ShortNumber.Precision - 1
        Number(i) = LongNumber.Digits(i) And ShortNumber.Digits(i)
    Next i
    
    For i = ShortNumber.Precision To LongNumber.Precision - 1
        Number(i) = LongNumber.Digits(i) And ExtDigit
    Next i
    
    If LongNumber.Sign = Negative Then
        Number(LongNumber.Precision) = &HFFFF And ExtDigit
    End If
    
    x.DetachNumber LongNumber
    y.DetachNumber ShortNumber
    
    Set AndBits = New BigInteger
    AndBits.InternalInit Number
End Function

''
' Peforms a bitwise OR operation, returning the result.
'
' @param x The left hand value of the OR operation.
' @param y The right hand value of the OR operation.
' @return A new BigInteger value containing the performed bitwise OR result.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function OrBits(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    ZeroIfNothing x
    ZeroIfNothing y
    
    If x.Sign = Sign.Zero Then
        Set OrBits = y
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set OrBits = x
        Exit Function
    End If
    
    Dim ExtDigit    As Integer
    Dim Number()    As Integer
    Dim LongNumber  As Number
    Dim ShortNumber As Number
    
    If x.Precision >= y.Precision Then
        x.AttachNumber LongNumber
        y.AttachNumber ShortNumber
    Else
        x.AttachNumber ShortNumber
        y.AttachNumber LongNumber
    End If
    
    If ShortNumber.Sign = Sign.Negative Then
        ExtDigit = &HFFFF
    End If
    
    ReDim Number(0 To LongNumber.Precision)
    
    Dim i As Long
    For i = 0 To ShortNumber.Precision - 1
        Number(i) = LongNumber.Digits(i) Or ShortNumber.Digits(i)
    Next i
    
    For i = ShortNumber.Precision To LongNumber.Precision - 1
        Number(i) = LongNumber.Digits(i) Or ExtDigit
    Next i
    
    If LongNumber.Sign = Negative Then
        Number(LongNumber.Precision) = &HFFFF Or ExtDigit
    End If
    
    x.DetachNumber LongNumber
    y.DetachNumber ShortNumber
    
    Set OrBits = New BigInteger
    OrBits.InternalInit Number
End Function

''
' Peforms a bitwise XOR operation, returning the result.
'
' @param x The left hand value of the XOR operation.
' @param y The right hand value of the XOR operation.
' @return A new BigInteger value containing the performed bitwise XOR result.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function XorBits(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    ZeroIfNothing x
    ZeroIfNothing y
    
    If x.Sign = Sign.Zero Then
        Set XorBits = y
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set XorBits = x
        Exit Function
    End If
    
    Dim ExtDigit    As Integer
    Dim Number()    As Integer
    Dim LongNumber  As Number
    Dim ShortNumber As Number
    
    If x.Precision >= y.Precision Then
        x.AttachNumber LongNumber
        y.AttachNumber ShortNumber
    Else
        x.AttachNumber ShortNumber
        y.AttachNumber LongNumber
    End If
    
    If ShortNumber.Sign = Sign.Negative Then
        ExtDigit = &HFFFF
    End If
    
    ReDim Number(0 To LongNumber.Precision)
    
    Dim i As Long
    For i = 0 To ShortNumber.Precision - 1
        Number(i) = LongNumber.Digits(i) Xor ShortNumber.Digits(i)
    Next i
    
    For i = ShortNumber.Precision To LongNumber.Precision - 1
        Number(i) = LongNumber.Digits(i) Xor ExtDigit
    Next i
    
    If LongNumber.Sign = Negative Then
        Number(LongNumber.Precision) = &HFFFF Xor ExtDigit
    End If
    
    x.DetachNumber LongNumber
    y.DetachNumber ShortNumber
    
    Set XorBits = New BigInteger
    XorBits.InternalInit Number
End Function

''
' Peforms a bitwise NOT of this BigInteger instance.
'
' @param x The BigInteger to perform the NOT operation on.
' @param value The value to perform the bitwise NOT on.
' @return A new BigInteger value containing the performed bitwise NOT result.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function NotBits(ByVal x As BigInteger) As BigInteger
    ZeroIfNothing x
    
    If x.Sign = Sign.Zero Then
        Set NotBits = mMinusOne
        Exit Function
    End If
    
    Dim xNumber As Number
    x.CopyNumber xNumber
        
    Dim i As Long
    For i = 0 To xNumber.Precision - 1
        xNumber.Digits(i) = Not xNumber.Digits(i)
    Next i
    
    Set NotBits = New BigInteger
    NotBits.InternalInitNumber xNumber
End Function

''
' Peforms a bitwise AND operation after NOTting the second operand, returning the result.
'
' @param x The left hand value of the operation.
' @param y The right hand value to be NOTted for the operation.
' @return A new BigInteger value containing the performed bitwise result.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function AndNotBits(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    Set AndNotBits = AndBits(x, NotBits(y))
End Function

''
' Raise a base value to the specified power.
'
' @param baseValue The value to be raised to a specified power.
' @param exponent The power to raise the value to.
' @return Returns the baseValue^exponent result.
' @remarks Values of Nothing are treated as zero.
'
Public Function Pow(ByVal BaseValue As BigInteger, ByVal Exponent As BigInteger) As BigInteger
    ZeroIfNothing BaseValue
    ZeroIfNothing Exponent

    If Exponent.Sign = Sign.Negative Then _
        Error.ArgumentOutOfRange "Exponent", ArgumentOutOfRange_NeedNonNegNum

    If Exponent.Sign = Sign.Zero Then
        Set Pow = mOne
        Exit Function
    ElseIf Exponent.IsOne Then
        Set Pow = BaseValue
        Exit Function
    End If

    Dim y       As BigInteger
    Dim z       As BigInteger
    Dim IsOdd   As Boolean
    
    Set y = mOne
    Set z = BaseValue
    
    Do
        IsOdd = Exponent.TestBit(0)
        Set Exponent = Exponent.ShiftRight(1)

        If IsOdd Then
            Set y = Multiply(y, z)
            
            If Exponent.IsZero Then
                Set Pow = y
                Exit Function
            End If
        End If
        
        Set z = Multiply(z, z)
    Loop
End Function

''
' Performs a modulo on a BigInteger value raised to a power. b = (x^y) mod z.
'
' @param baseValue The value to be raised to the power of the exponent.
' @param exponent The power to raise the baseValue to.
' @param m The modulus value to perform on the raised baseValue.
' @return Returns the modulus of the raised baseValue.
' <p>Values of Nothing are treated as zero.</p>
'
Public Function ModPow(ByVal BaseValue As BigInteger, ByVal Exponent As BigInteger, ByVal m As BigInteger) As BigInteger
    Set BaseValue = Pow(BaseValue, Exponent)
    DivRem BaseValue, m, ModPow
End Function

''
' Returns the mininum of the two BigInteger values.
'
' @param x One of the two values to be tested.
' @param y One of the two values to be tested.
' @return If x is less or equal to y, then x is returned, otherwise y is returned.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function Min(ByVal x As BigInteger, ByVal y As BigInteger) As BigInteger
    If Compare(x, y) <= 0 Then
        Set Min = x
    Else
        Set Min = y
    End If
End Function

''
' Returns the maximum of the two BigInteger values.
'
' @param x One of the two values to be tested.
' @param y One of the two values to be tested.
' @return If x is equal or greater than y, then x is returned, otherwise y is returned.
' <p>Values of Nothing will be treated as zero.</p>
'
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
'
Public Sub Randomize(Optional ByRef Value As Variant)
    VBA.Randomize Value
End Sub

''
' Sets the seeding of the Rnd function.
'
' @param seed The number to be used as the seed.
' @remarks This is provided because the VB Rnd function runs
' independently for each component.
'
Public Sub SeedRnd(Optional ByRef seed As Variant)
    VBA.Rnd seed
End Sub

''
' Generates a random BigInteger number of the specified size.
'
' @param size The size of the number in bytes.
' @param IsNegative Indicates if a negative random number should be generated.
' @return A new BigInteger of a random value.
'
Public Function Rnd(ByVal Size As Long, Optional ByVal IsNegative As Boolean = False) As BigInteger
    If Size < 0 Then _
        Error.ArgumentOutOfRange "Size", ArgumentOutOfRange_NeedNonNegNum
    
    If Size = 0 Then
        Set Rnd = mZero
        Exit Function
    End If
    
    Dim WordCount As Long
    WordCount = Size \ 2
    
    Dim Ints() As Integer
    ReDim Ints(0 To WordCount)
    
    Dim i As Long
    For i = 0 To WordCount - 1
        Ints(i) = Int(VBA.Rnd * 65536) - 32768
    Next i
    
    If Size And 1 Then
        Ints(WordCount) = VBA.Rnd * &H100
    End If
    
    If IsNegative Then
        ApplyTwosComplement Ints
    End If
    
    Set Rnd = New BigInteger
    Rnd.InternalInit Ints
End Function

''
' Performs a left shift operation on the BigInteger instance.
'
' @param x The BigInteger to shift left.
' @param bitCount The number of bits the value is to be shifted left.
' @return Returns a new BigInteger instance containing the shifted value.
' @remarks If bitCount is negative, then a right shift is performed.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function ShiftLeft(ByVal x As BigInteger, ByVal BitCount As Long) As BigInteger
    ZeroIfNothing x
    Set ShiftLeft = x.ShiftLeft(BitCount)
End Function

''
' Performs a bit shifting operation to the the right.
'
' @param x The BigInteger to shift right.
' @param bitCount The number of bits the BigInteger value is to be shifted.
' @return Returns a new BigInteger instance containing the bit shifted value.
' @remarks If bitCount is negative, then a left shift is performed.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function ShiftRight(ByVal x As BigInteger, ByVal BitCount As Long) As BigInteger
    ZeroIfNothing x
    Set ShiftRight = x.ShiftRight(BitCount)
End Function

''
' Returns the integer square root of a BigInteger.
'
' @param x The value to calculate the integer square root for.
' @Return The integer square root of the value.
'
Public Function SquareRoot(ByVal x As BigInteger) As BigInteger
    ZeroIfNothing x
    
    If x.Sign = Sign.Zero Then
        Set SquareRoot = mZero
        Exit Function
    ElseIf x.Sign = Negative Then
        Set x = x.Negate
    End If
    
    Dim Estimate As BigInteger
    Dim n As BigInteger
    
    Set Estimate = x.ShiftRight(1).Add(mOne)
    Set n = x.Divide(Estimate).Add(Estimate).ShiftRight(1)
    
    Do While Compare(n, Estimate) < 0
        Set Estimate = n
        Set n = x.Divide(Estimate).Add(Estimate).ShiftRight(1)
    Loop
    
    Set SquareRoot = Estimate
End Function

''
' Calculates the factorial of an integer value.
'
' @param Value The value to calculate the factorial of.
' @return A BigInteger value of the factorial.
'
Public Function Factorial(ByVal Value As BigInteger) As BigInteger
    Dim Result As BigInteger
    
    ZeroIfNothing Value
    
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
Private Sub ZeroIfNothing(ByRef RefValue As BigInteger)
    If RefValue Is Nothing Then
        Set RefValue = mZero
    End If
End Sub

Private Function TryParseDecimal(ByRef s As String, ByVal Style As NumberStyles, ByVal Info As NumberFormatInfo, ByRef n As Number) As Boolean
    Const NotFound          As Long = 0
    Const FoundBefore       As Long = 1
    Const FoundAfter        As Long = 2
    Const FoundDuring       As Long = 3
    Dim Chars()             As Integer
    Dim i                   As Long
    Dim Ch                  As Long
    Dim InWhiteSpace        As Boolean
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
    
    NegativeSign = AscW(Info.NegativeSign)
    PositiveSign = AscW(Info.PositiveSign)
    CurrencySymbol = AscW(Info.CurrencySymbol)
    DecimalPoint = AscW(Info.NumberDecimalSeparator)
    CurrencyPoint = AscW(Info.CurrencyDecimalSeparator)
    Thousands = AscW(Info.NumberGroupSeparator)
    CurrencyThousands = AscW(Info.CurrencyGroupSeparator)
    Chars = AllocChars(s)
    ReDim n.Digits(0 To Len(s) \ 4)
    n.Sign = 1
    
    For i = 0 To UBound(Chars)
        Ch = Chars(i)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                If (FoundCurrency = FoundAfter) Or (FoundSign = FoundAfter) Or InWhiteSpace Then
                    GoTo Finally
                End If
                
                If Not FoundDecimal Then
                    SingleInPlaceMultiply n, 10
                    SingleInPlaceAdd n, Ch - vbZeroChar
                ElseIf Ch <> vbZeroChar Then
                    GoTo Finally
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
                
                    InWhiteSpace = FoundSign
                Else
                    InWhiteSpace = True
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
            Case Else
                GoTo Finally
        End Select
    Next
    
    If InWhiteSpace Then
        If (Style And NumberStyles.AllowTrailingWhite) = 0 Then
            GoTo Finally
        End If
    End If
    
    If NeedsClosingParen Then
        GoTo Finally
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

Private Function TryParseHex(ByRef s As String, ByRef n As Number) As Boolean
    Dim Chars() As Integer
    Dim ub      As Long
    
    If Len(s) < 3 Then
        Exit Function
    End If
    
    Chars = AllocChars(s)
    ub = (Len(s) - 2) \ 4 + 1
    ReDim n.Digits(0 To ub)
    
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    j = 1
    For i = Len(s) - 1 To 2 Step -1
        Dim Ch As Long
        Select Case Chars(i)
            Case vbZeroChar To vbNineChar
                Ch = Chars(i) - vbZeroChar
            Case 65 To 70
                Ch = Chars(i) - 55
            Case 97 To 102
                Ch = Chars(i) - 87
            Case Else
                GoTo Finally
        End Select
        
        k = k + Ch * j
        j = j * 16
        
        If j = &H10000 Then
            n.Digits(c) = AsWord(k)
            k = 0
            j = 1
            c = c + 1
        End If
    Next i
    
    If k Then
        n.Digits(c) = AsWord(k)
        c = c + 1
    End If
    
    n.Precision = c
    TryParseHex = True
    
Finally:
    FreeChars Chars
End Function

Private Sub DivideByZeroError()
    Call Err.Raise(11)
End Sub

Private Function InternalDivide(ByRef Dividend As BigInteger, ByRef Divisor As BigInteger, ByRef remainder As BigInteger, ByVal IncludeRemainder As Boolean) As BigInteger
    ZeroIfNothing Dividend
    ZeroIfNothing Divisor
    
    If Divisor.Sign = Sign.Zero Then
        DivideByZeroError
    ElseIf Dividend.Sign = Sign.Zero Then
        Set InternalDivide = mZero
        If IncludeRemainder Then
            Set remainder = mZero
        End If
        Exit Function
    ElseIf Divisor.IsOne Then
        Set InternalDivide = Dividend
        If IncludeRemainder Then
            Set remainder = mZero
        End If
        Exit Function
    ElseIf Divisor.IsMinusOne Then
        Set InternalDivide = Dividend.Negate
        If IncludeRemainder Then
            Set remainder = mZero
        End If
        Exit Function
    End If
    
    Dim xNumber As Number
    Dim yNumber As Number
    
    Dividend.CopyNumber xNumber
    Divisor.CopyNumber yNumber
      
    If Dividend.Sign = Sign.Negative Then
        Mathematics.Negate xNumber
    End If
    If Divisor.Sign = Sign.Negative Then
        Mathematics.Negate yNumber
    End If
    
    Dim QuotientNumber()    As Integer
    Dim RemainderNumber()   As Integer
    
    If yNumber.Precision > 1 Then
        QuotientNumber = GradeSchoolDivide(xNumber, yNumber, RemainderNumber, IncludeRemainder)
    Else
        Dim R As Long
        QuotientNumber = SinglePlaceDivide(xNumber.Digits, xNumber.Precision, yNumber.Digits(0), R)
        
        If IncludeRemainder Then
            ReDim RemainderNumber(0)
            RemainderNumber(0) = R
        End If
    End If
    
    If Dividend.Sign <> Divisor.Sign Then
        ApplyTwosComplement QuotientNumber
        
        If IncludeRemainder Then
            ApplyTwosComplement RemainderNumber
        End If
    End If
    
    If IncludeRemainder Then
        Set remainder = New BigInteger
        remainder.InternalInit RemainderNumber
                
        If remainder.Sign = Negative Then
            Set remainder = Negate(remainder)
        End If
    End If
    
    Set InternalDivide = New BigInteger
    InternalDivide.InternalInit QuotientNumber
End Function


