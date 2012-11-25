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

Private mChars()    As Integer
Private mCharsSA    As SafeArray1d
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
    Dim xNumber As Number: Call x.AttachNumber(xNumber)
    Dim yNumber As Number: Call y.AttachNumber(yNumber)
    
    Dim Result As Long
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
' @param s The string of characters to attempt to be parsed.
' @param b The returning BigInteger object if the string was successfully parsed.
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
'
Public Function TryParse(ByVal s As String, ByRef b As BigInteger) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    
    Dim StartIndex As Long
    Dim n As Number
    
    If Left$(s, 1) = "-" Then
        StartIndex = 1
        n.Sign = Negative
    End If

    Select Case Mid$(s, StartIndex + 1, 2)
        Case "&h", "&H", "0X", "0x":    TryParse = TryParseHex(s, StartIndex, n)
        Case "0b", "0B":                TryParse = TryParseBinary(s, StartIndex, n)
        Case Else:                      TryParse = TryParseDecimal(s, StartIndex, n)
    End Select
    
    If TryParse Then
        If n.Sign = Negative Then Call modBigMath.Negate(n)
        
        Set b = New BigInteger
        Call b.InternalInitNumber(n)
    End If
End Function

''
' Attemps to parse a string of characters of the specified base type.
'
' @param s The string of characters to attempt to be parsed.
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
Public Function Parse(ByVal s As String) As BigInteger
    If Not TryParse(s, Parse) Then
        Throw Cor.NewArgumentException("Number could not be parsed.", "s")
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
    If x Is Nothing Then Set x = mZero
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
    If x Is Nothing Then Set x = mZero
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
    If x Is Nothing Then Set x = mZero
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
    If x.Sign = Sign.Zero Then
        Set Add = y
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set Add = x
        Exit Function
    End If
    
    Dim u As Number: Call x.AttachNumber(u)
    Dim v As Number: Call y.AttachNumber(v)
    
    Dim sum() As Integer
    sum = GradeSchoolAdd(u, v)
    
    Call x.DetachNumber(u)
    Call y.DetachNumber(v)
    
    Set Add = New BigInteger
    Call Add.InternalInit(sum)
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
    If x.Sign = Sign.Zero Then
        Set Subtract = y.Negate
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set Subtract = x
        Exit Function
    End If
    
    Dim u As Number: Call x.AttachNumber(u)
    Dim v As Number: Call y.AttachNumber(v)

    Dim difference() As Integer
    difference = GradeSchoolSubtract(u, v)

    Call x.DetachNumber(u)
    Call y.DetachNumber(v)
    
    Set Subtract = New BigInteger
    Call Subtract.InternalInit(difference)
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
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
        Call x.CopyNumber(u)
        Call modBigMath.Negate(u)
    Else
        Call x.AttachNumber(u)
    End If

    If y.Sign = Sign.Negative Then
        Call y.CopyNumber(v)
        Call modBigMath.Negate(v)
    Else
        Call y.AttachNumber(v)
    End If
    
    Dim product() As Integer
    product = GradeSchoolMultiply(u, v)
    
    If x.Sign = Sign.Positive Then Call x.DetachNumber(u)
    If y.Sign = Sign.Positive Then Call y.DetachNumber(v)
    If x.Sign <> y.Sign Then Call ApplyTwosComplement(product)
        
    Set Multiply = New BigInteger
    Call Multiply.InternalInit(product)
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
Public Function Divide(ByVal dividend As BigInteger, ByVal divisor As BigInteger) As BigInteger
    Set Divide = InternalDivide(dividend, divisor, Nothing, False)
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
'
Public Function DivRem(ByVal dividend As BigInteger, ByVal divisor As BigInteger, ByRef remainder As BigInteger) As BigInteger
    Set DivRem = InternalDivide(dividend, divisor, remainder, True)
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
'
Public Function Modulus(ByVal dividend As BigInteger, ByVal divisor As BigInteger) As BigInteger
    Call InternalDivide(dividend, divisor, Modulus, True)
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
    If x.Sign = Sign.Zero And y.Sign = Sign.Zero Then
        Set GreatestCommonDivisor = mZero
        Exit Function
    End If
    
    Set x = x.Absolute
    Set y = y.Absolute
    
    Dim r As BigInteger
    Do While y.Sign <> Sign.Zero
        Set r = Modulus(x, y)
        Set x = y
        Set y = r
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
Public Function TestBit(ByVal x As BigInteger, ByVal bit As Long) As Boolean
    If x Is Nothing Then Set x = mZero
    TestBit = x.TestBit(bit)
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
Public Function SetBit(ByVal x As BigInteger, ByVal bit As Long) As BigInteger
    If x Is Nothing Then Set x = mZero
    SetBit = x.SetBit(bit)
End Function

''
' Clears the specified bit of the BigInteger instance to 0.
'
' @param x The BigInteger value to clear the bit on.
' @param bit The bit to be set to 0.
' @return A new BigInteger value with the bit set to 0.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function ClearBit(ByVal x As BigInteger, ByVal bit As Long) As BigInteger
    If x Is Nothing Then Set x = mZero
    Set ClearBit = x.ClearBit(bit)
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
Public Function FlipBit(ByVal x As BigInteger, ByVal bit As Long) As BigInteger
    If x Is Nothing Then Set x = mZero
    Set FlipBit = x.FlipBit(bit)
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
    If x.Sign = Sign.Zero Or y.Sign = Sign.Zero Then
        Set AndBits = mZero
        Exit Function
    End If
    
    Dim extDigit    As Integer
    Dim Number()    As Integer
    Dim longNumber  As Number
    Dim shortNumber As Number
    
    If x.Precision >= y.Precision Then
        Call x.AttachNumber(longNumber)
        Call y.AttachNumber(shortNumber)
    Else
        Call x.AttachNumber(shortNumber)
        Call y.AttachNumber(longNumber)
    End If
    If shortNumber.Sign = Sign.Negative Then extDigit = &HFFFF
    ReDim Number(0 To longNumber.Precision)
    
    Dim i As Long
    For i = 0 To shortNumber.Precision - 1
        Number(i) = longNumber.Digits(i) And shortNumber.Digits(i)
    Next i
    
    For i = shortNumber.Precision To longNumber.Precision - 1
        Number(i) = longNumber.Digits(i) And extDigit
    Next i
    
    If longNumber.Sign = Negative Then Number(longNumber.Precision) = &HFFFF And extDigit
    
    Call x.DetachNumber(longNumber)
    Call y.DetachNumber(shortNumber)
    
    Set AndBits = New BigInteger
    Call AndBits.InternalInit(Number)
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
    If x.Sign = Sign.Zero Then
        Set OrBits = y
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set OrBits = x
        Exit Function
    End If
    
    Dim extDigit    As Integer
    Dim Number()    As Integer
    Dim longNumber  As Number
    Dim shortNumber As Number
    
    If x.Precision >= y.Precision Then
        Call x.AttachNumber(longNumber)
        Call y.AttachNumber(shortNumber)
    Else
        Call x.AttachNumber(shortNumber)
        Call y.AttachNumber(longNumber)
    End If
    If shortNumber.Sign = Sign.Negative Then extDigit = &HFFFF
    ReDim Number(0 To longNumber.Precision)
    
    Dim i As Long
    For i = 0 To shortNumber.Precision - 1
        Number(i) = longNumber.Digits(i) Or shortNumber.Digits(i)
    Next i
    
    For i = shortNumber.Precision To longNumber.Precision - 1
        Number(i) = longNumber.Digits(i) Or extDigit
    Next i
    
    If longNumber.Sign = Negative Then Number(longNumber.Precision) = &HFFFF Or extDigit
    
    Call x.DetachNumber(longNumber)
    Call y.DetachNumber(shortNumber)
    
    Set OrBits = New BigInteger
    Call OrBits.InternalInit(Number)
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
    If x Is Nothing Then Set x = mZero
    If y Is Nothing Then Set y = mZero
    
    If x.Sign = Sign.Zero Then
        Set XorBits = y
        Exit Function
    ElseIf y.Sign = Sign.Zero Then
        Set XorBits = x
        Exit Function
    End If
    
    Dim extDigit    As Integer
    Dim Number()    As Integer
    Dim longNumber  As Number
    Dim shortNumber As Number
    
    If x.Precision >= y.Precision Then
        Call x.AttachNumber(longNumber)
        Call y.AttachNumber(shortNumber)
    Else
        Call x.AttachNumber(shortNumber)
        Call y.AttachNumber(longNumber)
    End If
    If shortNumber.Sign = Sign.Negative Then extDigit = &HFFFF
    ReDim Number(0 To longNumber.Precision)
    
    Dim i As Long
    For i = 0 To shortNumber.Precision - 1
        Number(i) = longNumber.Digits(i) Xor shortNumber.Digits(i)
    Next i
    
    For i = shortNumber.Precision To longNumber.Precision - 1
        Number(i) = longNumber.Digits(i) Xor extDigit
    Next i
    
    If longNumber.Sign = Negative Then Number(longNumber.Precision) = &HFFFF Xor extDigit
    
    Call x.DetachNumber(longNumber)
    Call y.DetachNumber(shortNumber)
    
    Set XorBits = New BigInteger
    Call XorBits.InternalInit(Number)
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
    If x Is Nothing Then Set x = mZero
    
    If x.Sign = Sign.Zero Then
        Set NotBits = mMinusOne
        Exit Function
    End If
    
    Dim xNumber As Number
    Call x.CopyNumber(xNumber)
        
    Dim i As Long
    For i = 0 To xNumber.Precision - 1
        xNumber.Digits(i) = Not xNumber.Digits(i)
    Next i
    
    Set NotBits = New BigInteger
    Call NotBits.InternalInitNumber(xNumber)
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
Public Function Pow(ByVal baseValue As BigInteger, ByVal exponent As BigInteger) As BigInteger
    If baseValue Is Nothing Then Set baseValue = mZero
    If exponent Is Nothing Then Set exponent = mZero

    If exponent.Sign = Sign.Negative Then _
        Throw Cor.NewArgumentException("Exponent cannot be negative.", "exponent")

    If exponent.Sign = Sign.Zero Then
        Set Pow = mOne
        Exit Function
    ElseIf exponent.IsOne Then
        Set Pow = baseValue
        Exit Function
    End If

    Dim y As BigInteger: Set y = mOne
    Dim z As BigInteger: Set z = baseValue

    Dim isOdd As Boolean
    Do
        isOdd = exponent.TestBit(0)
        Set exponent = exponent.ShiftRight(1)

        If isOdd Then
            Set y = Multiply(y, z)
            If exponent.Sign = Sign.Zero Then
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
Public Function ModPow(ByVal baseValue As BigInteger, ByVal exponent As BigInteger, ByVal m As BigInteger) As BigInteger
    Set baseValue = Pow(baseValue, exponent)
    Call DivRem(baseValue, m, ModPow)
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
' @param num The number used to randomize the Rnd method.
' @remarks This is provided because the VB Rnd function runs
' independently for each component.
'
Public Sub Randomize(Optional ByVal num As Variant)
    Call VBA.Randomize(num)
End Sub

''
' Sets the seeding of the Rnd function.
'
' @param seed The number to be used as the seed.
' @remarks This is provided because the VB Rnd function runs
' independently for each component.
'
Public Sub SeedRnd(Optional ByVal seed As Variant)
    Call VBA.Rnd(seed)
End Sub

''
' Generates a random BigInteger number of the specified size.
'
' @param size The size of the number in bytes.
' @param isNegative Indicates if a negative random number should be generated.
' @return A new BigInteger of a random value.
'
Public Function Rnd(ByVal Size As Long, Optional ByVal isNegative As Boolean = False) As BigInteger
    If Size < 0 Then Call Err.Raise(5, , "size cannot be negative")
    
    If Size = 0 Then
        Set Rnd = mZero
        Exit Function
    End If
    
    Dim wordCount As Long
    wordCount = Size \ 2
    
    Dim ints() As Integer
    ReDim ints(0 To wordCount)
    
    Dim i As Long
    For i = 0 To wordCount - 1
        ints(i) = Int(VBA.Rnd * 65536) - 32768
    Next i
    
    If Size And 1 Then
        ints(wordCount) = VBA.Rnd * &H100
    End If
    
    If isNegative Then Call ApplyTwosComplement(ints)
    
    Set Rnd = New BigInteger
    Call Rnd.InternalInit(ints)
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
Public Function ShiftLeft(ByVal x As BigInteger, ByVal bitCount As Long) As BigInteger
    If x Is Nothing Then Set x = mZero
    Set ShiftLeft = x.ShiftLeft(bitCount)
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
Public Function ShiftRight(ByVal x As BigInteger, ByVal bitCount As Long) As BigInteger
    If x Is Nothing Then Set x = mZero
    Set ShiftRight = x.ShiftRight(bitCount)
End Function

''
' Returns the integer square root of a BigInteger.
'
' @param x The value to calculate the integer square root for.
' @Return The integer square root of the value.
'
Public Function SquareRoot(ByVal x As BigInteger) As BigInteger
    If x Is Nothing Then Set x = mZero
    
    If x.Sign = Sign.Zero Then
        Set SquareRoot = mZero
        Exit Function
    ElseIf x.Sign = Negative Then
        Set x = x.Negate
    End If
    
    Dim estimate As BigInteger
    Set estimate = x.ShiftRight(1).Add(mOne)
    
    Dim n As BigInteger
    Set n = x.Divide(estimate).Add(estimate).ShiftRight(1)
    
    Do While Compare(n, estimate) < 0
        Set estimate = n
        Set n = x.Divide(estimate).Add(estimate).ShiftRight(1)
    Loop
    
    Set SquareRoot = estimate
End Function

''
' Calculates the factorial of an integer value.
'
' @param x The value to calculate the factorial of.
' @return A BigInteger value of the factorial.
'
Public Function Factorial(ByVal x As BigInteger) As BigInteger
    If x Is Nothing Then Set x = mZero
    
    Select Case x.Sign
        Case Sign.Zero
            Set Factorial = mZero
            Exit Function
        
        Case Sign.Negative
            Set x = x.Negate
    
    End Select
    
    Dim Result As BigInteger
    Set Result = mOne
    
    Do While Not (x.Sign = Sign.Zero)
        Set Result = Result.Multiply(x)
        Set x = x.Subtract(mOne)
    Loop
    
    Set Factorial = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TryParseDecimal(ByRef s As String, ByVal StartIndex As Long, ByRef n As Number) As Boolean
    If Len(s) - StartIndex < 1 Then Exit Function
    
    Call HookString(s)
    Dim ub As Long: ub = Len(s) \ 4 + 1
    ReDim n.Digits(0 To ub)
    
    Dim i As Long
    For i = StartIndex To mCharsSA.cElements - 1
        Select Case mChars(i)
            Case 48 To 57
                Call SingleInPlaceMultiply(n, 10)
                Call SingleInPlaceAdd(n, mChars(i) - 48)
            Case Else
                Exit Function
        End Select
    Next i
    
    n.Precision = n.Precision + 1   ' add one additional zero.
    Call UnhookString
    
    TryParseDecimal = True
End Function

Private Function TryParseHex(ByRef s As String, ByVal StartIndex As Long, ByRef n As Number) As Boolean
    If Len(s) - StartIndex < 3 Then Exit Function
    
    Call HookString(s)
    Dim ub As Long: ub = (Len(s) - 2) \ 4 + 1
    ReDim n.Digits(0 To ub)
    
    Dim c As Long
    Dim i As Long
    Dim j As Long: j = 1
    Dim k As Long
    For i = Len(s) - 1 To 2 + StartIndex Step -1
        Dim ch As Long
        Select Case mChars(i)
            Case 48 To 57:  ch = mChars(i) - 48
            Case 65 To 70:  ch = mChars(i) - 55
            Case 97 To 102: ch = mChars(i) - 87
            Case Else
                Exit Function
        End Select
        
        k = k + ch * j
        j = j * 16
        
        If j = &H10000 Then
            n.Digits(c) = GetInt(k)
            k = 0
            j = 1
            c = c + 1
        End If
    Next i
    
    If k Then
        n.Digits(c) = GetInt(k)
        c = c + 1
    End If
    
    n.Precision = c
    Call UnhookString
    
    TryParseHex = True
End Function

Private Function TryParseBinary(ByRef s As String, ByVal StartIndex As Long, ByRef n As Number) As Boolean
    If Len(s) - StartIndex < 3 Then Exit Function
    
    Call HookString(s)
    
    Dim ub As Long: ub = (Len(s) - 2) \ 16 + 1
    If ub = 0 Then ub = 1
    ReDim n.Digits(0 To ub)
    
    Dim c As Long
    Dim i As Long
    Dim j As Long: j = 1
    Dim k As Long
    For i = Len(s) - 1 To 2 + StartIndex Step -1
        Dim ch As Long
        Select Case mChars(i)
            Case 48, 49: ch = mChars(i) - 48
            Case Else
                Exit Function
        End Select
        
        k = k + ch * j
        j = j * 2
        
        If j = &H10000 Then
            n.Digits(c) = GetInt(k)
            j = 1
            c = c + 1
            k = 0
        End If
    Next i
    
    If k Then
        n.Digits(c) = GetInt(k)
        c = c + 1
    End If
    
    n.Precision = c
    Call UnhookString
    
    TryParseBinary = True
End Function

Private Sub DivideByZeroError()
    Call Err.Raise(11)
End Sub

Private Function InternalDivide(ByRef dividend As BigInteger, ByRef divisor As BigInteger, ByRef remainder As BigInteger, ByVal includeRemainder As Boolean) As BigInteger
    If dividend Is Nothing Then Set dividend = mZero
    If divisor Is Nothing Then Set divisor = mZero
    
    If divisor.Sign = Sign.Zero Then
        Call DivideByZeroError
    ElseIf dividend.Sign = Sign.Zero Then
        Set InternalDivide = mZero
        If includeRemainder Then Set remainder = mZero
        Exit Function
    ElseIf divisor.IsOne Then
        Set InternalDivide = dividend
        If includeRemainder Then Set remainder = mZero
        Exit Function
    ElseIf divisor.IsMinusOne Then
        Set InternalDivide = dividend.Negate
        If includeRemainder Then Set remainder = mZero
        Exit Function
    End If
    
    Dim xNumber As Number: Call dividend.CopyNumber(xNumber)
    Dim yNumber As Number: Call divisor.CopyNumber(yNumber)
    
    If dividend.Sign = Sign.Negative Then Call modBigMath.Negate(xNumber)
    If divisor.Sign = Sign.Negative Then Call modBigMath.Negate(yNumber)
    
    Dim quotientNumber()    As Integer
    Dim remainderNumber()   As Integer
    
    If yNumber.Precision > 1 Then
        quotientNumber = GradeSchoolDivide(xNumber, yNumber, remainderNumber, includeRemainder)
    Else
        Dim r As Long
        quotientNumber = SinglePlaceDivide(xNumber.Digits, xNumber.Precision, yNumber.Digits(0), r)
        
        If includeRemainder Then
            ReDim remainderNumber(0)
            remainderNumber(0) = r
        End If
    End If
    
    If dividend.Sign <> divisor.Sign Then
        Call ApplyTwosComplement(quotientNumber)
        
        If includeRemainder Then
            Call ApplyTwosComplement(remainderNumber)
        End If
    End If
    
    If includeRemainder Then
        Set remainder = New BigInteger
        Call remainder.InternalInit(remainderNumber)
    End If
    
    Set InternalDivide = New BigInteger
    Call InternalDivide.InternalInit(quotientNumber)
End Function

Private Sub HookString(ByRef s As String)
    mCharsSA.pvData = StrPtr(s)
    mCharsSA.cElements = Len(s)
    mCharsSA.cDims = 1
    mCharsSA.cbElements = 2
    SAPtr(mChars) = VarPtr(mCharsSA)
End Sub

Private Sub UnhookString()
    SAPtr(mChars) = vbNullPtr
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mZero = Cor.NewBigInteger(0)
    Set mOne = Cor.NewBigInteger(1)
    Set mMinusOne = Cor.NewBigInteger(-1)
    Set mTen = Cor.NewBigInteger(10)
End Sub

