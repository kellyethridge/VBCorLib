VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: BigInteger
'

''
' Represents an unlimited size signed integer value.
'
' @remarks The BigInteger can represent a positive or negative value limited in size
' by only by memory. Negative values are represented in a
' Twos Complement fashion mimicking the way the hardware and VB uses signed integer values.
' <p>The value is represented internally as an array of Integers. This was selected over
' Bytes because of the speed increase with working with multiple bytes simultaneously.
' However, Integers are never revealed outside of the class, only byte arrays.
'
Option Explicit
Implements IObject
Implements IComparable
Implements IFormattable

Private Const BytesPerDigit As Long = vbSizeOfInteger
Private Const HEX_CHARS As String = "0123456789abcdef0123456789ABCDEF"
Private Const LOWER_HEX As Long = 1
Private Const UPPER_HEX As Long = 17

Private mNumber As BigNumber

Private Type Digit
    Index As Long
    Value As Long
End Type


''
' Returns the sign of the number this instance represents.
'
' @return The value of type Sign representing the sign of the number.
' @remarks Returns -1 for negative, 0 for zero, 1 for positive.
Public Property Get Sign() As Sign
    Sign = mNumber.Sign
End Property

''
' Returns if the value is equal to zero.
'
' @return <b>True</b> if the value is zero, otherwise <b>False</b>.
Public Property Get IsZero() As Boolean
    IsZero = mNumber.Sign = Zero
End Property

''
' Returns if the value is equal to one.
'
' @return <b>True</b> if the value is one, otherwise <b>False</b>.
Public Property Get IsOne() As Boolean
    If mNumber.Sign = Positive And mNumber.Precision = 1 Then
        IsOne = mNumber.Digits(0) = 1
    End If
End Property

''
' Returns the number of bytes being used to represent the number.
'
' @return The number of bytes being used to represent the number.
' @remarks There may be additional bytes allocated, but not considered
' as part of the number. Those bytes are not represented in the precision.
' <p>Since a negative value can have an infinite number of leading &HFFFF values,
' only 1 (2 bytes) is counted if it is needed.</p>
Public Property Get Precision() As Long
    Const BytesPerDigit As Long = 2
    Dim Result As Long
    
    Result = mNumber.Precision * BytesPerDigit
    
    If Result > 0 Then
        Dim LastDigitIsPartial As Boolean
        
        LastDigitIsPartial = (mNumber.Digits(mNumber.Precision - 1) And &HFF00) = 0
        
        If LastDigitIsPartial Then
            Result = Result - 1
        End If
    End If
    
    Precision = Result
End Property

''
' Returns a byte array representing the number.
'
' @return A byte array the same length as Precision that represents the number.
' @remarks To represent a byte array for a negative value, the highest bit must be
' set. To ensure this, a pair of bytes with the value of &HFF may be part of the return
' value.
'
Public Function ToByteArray() As Byte()
    Dim Result()    As Byte
    Dim MaxIndex    As Long
    Dim Size        As Long
    Dim CopyCount   As Long
    Dim SetHighBits As Boolean
    
    If mNumber.Sign = 0 Then
        ReDim Result(0)
    Else
        MaxIndex = mNumber.Precision - 1
        Size = (MaxIndex + 1) * 2
        CopyCount = Size
        
        If mNumber.Sign < 0 Then
            ' if the negative bits extend through the upper 8 bits
            ' of the last digit we don't want to copy those 8 bits.
            Select Case mNumber.Digits(MaxIndex) And &HFF80
                Case &HFF80, &H80
                    ' we don't need the leading &hFF or &h0 so scale
                    ' back the result and number of bytes to be copied.
                    Size = Size - 1
                    CopyCount = CopyCount - 1
                Case Is >= 0
                    ' we are negative but have not high bit set so
                    ' we are going to add one byte to the output and
                    ' set it to &hFF.
                    Size = Size + 1
                    SetHighBits = True
            End Select
        Else
            If mNumber.Digits(MaxIndex) And &H8000 Then
                Size = Size + 1
            ElseIf (mNumber.Digits(MaxIndex) And &HFF80) = 0 Then
                Size = Size - 1
                CopyCount = CopyCount - 1
            End If
        End If
    
        ReDim Result(0 To Size - 1)
        CopyMemory Result(0), mNumber.Digits(0), CopyCount
        
        If SetHighBits Then
            Result(Size - 1) = &HFF
        End If
    End If
        
    ToByteArray = Result
End Function

''
' Returns a human readable string of characters representing the number.
'
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
Public Function ToString(Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim Specifier   As String
    Dim Count       As Long
    Dim FormatInfo  As NumberFormatInfo
    
    Count = ParseFormatSpecifier(Format, Specifier)
    Set FormatInfo = NumberFormatInfo.GetInstance(Provider)
    
    Select Case Specifier
        Case "x":       ToString = ToHexString(Count, LOWER_HEX)
        Case "X":       ToString = ToHexString(Count, UPPER_HEX)
        Case "b", "B":  ToString = ToBinaryString(Count)
        Case "d", "D":  ToString = ToDecimalString(Count, FormatInfo)
    End Select
End Function

''
' Compares two BigInteger instance to each other returning their relation.
'
' @param other The right handed value to compare this instance against.
' @return Returns the relation between this instance and the other number.<br>
' @remarks If Nothing is passed in, it is treated as zero.<br>
' If this instance < other then -1 is returned<br>
' If this instance = other then 0 is returned<br>
' If this instance > other then 1 is returned.<br>
'
Public Function CompareTo(ByVal Other As BigInteger) As Long
    CompareTo = BigInteger.Compare(Me, Other)
End Function

''
' Determines if this BigInteger instance is equal to the other value.
'
' @param other The other value to determine if this instance is equal to.
' @return Returns True if this object and the other value are equal, False otherwise.
' @remarks The other value accepts a String, Number, Byte Array, or BigInteger
' object. The BInt function is called to convert the other value to
' a BigInteger object before comparison.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As BigInteger
    
    On Error GoTo Catch
    Set Other = ValueOrZero(Value)
    Equals = Other.EqualNumbers(mNumber)
    
Catch:
End Function

Friend Function EqualNumbers(ByRef Other As BigNumber) As Boolean
    EqualNumbers = BigNumberMath.Equals(mNumber, Other)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @IObject
'
Public Function GetHashCode() As Long
    Dim Result          As Long
    Dim DoubleDigits()  As Long
    Dim Count           As Long
    
    Result = mNumber.Sign Xor mNumber.Precision
    SAPtr(DoubleDigits) = SAPtr(mNumber.Digits)
    Count = mNumber.Precision \ 2
    
    Dim i As Long
    For i = 0 To Count - 1
        Result = Result Xor DoubleDigits(i)
    Next
    
    If mNumber.Precision And 1 Then
        Result = Result Xor mNumber.Digits(mNumber.Precision - 1)
    End If
    
    SAPtr(DoubleDigits) = vbNullPtr
    GetHashCode = Result
End Function

''
' Returns the opposite signed value of this BigInteger instance.
'
' @return Returns a BigInteger instance containing the opposite signed value.
' @remarks Negative values are represented as a Twos Complement number, unlike
' other systems used to represent a negative value. Any negation requires
' the Twos Complement algorithm to be applied to the number. This keeps in step
' with how users of VB already deal with negative numbers, helping to ease
' the transition to using BigInteger values.
'
Public Function Negate() As BigInteger
    If mNumber.Sign = Zero Then
        Set Negate = Me
    Else
        Dim Copy As BigNumber
        
        Copy = mNumber
        BigNumberMath.Negate Copy
        Set Negate = New BigInteger
        Negate.InternalInitNumber2 Copy
    End If
End Function

''
' Returns the absolute (positive or zero) value of this instance.
'
' @return Returns a positive (or zero) value of the current BigInteger instance.
' @remarks If the value is already positive or zero, then the same object instance
' is returned. If the value is negative, then the negated value is returned.
'
Public Function Absolute() As BigInteger
    If mNumber.Sign = Negative Then
        Set Absolute = Negate
    Else
        Set Absolute = Me
    End If
End Function

''
' Adds a value to this instance of BigInteger.
'
' @param Value A numeric representation of the value to be added to this instance.
' @return A new instance of BigInteger containing the summed value.
' @remarks This method accepts a BigInteger, one of VB's numeric values or
' a string to be parsed.
' <pre>
' Dim b As BigInteger
' Set b = BInt(100)
' Set b = b.Add(BInt(150))
' Debug.Print b.ToString '' 250
' </pre>
'
Public Function Add(ByVal Value As BigInteger) As BigInteger
    Set Add = BigInteger.Add(Me, Value)
End Function

''
' Subtracts a value to this instance of BigInteger.
'
' @param Value A numeric representation of the value to be subtracted to this instance.
' @return A new instance of BigInteger containing the differenced value.
' @remarks This method accepts a BigInteger, one of VB's numeric values or
' a string to be parsed.
' <pre>
' Dim b As BigInteger
' Set b = BInt(100)
' Set b = b.Subtract(BInt(150))
' Debug.Print b.ToString '' -50
' </pre>
'
Public Function Subtract(ByVal Value As BigInteger) As BigInteger
    Set Subtract = BigInteger.Subtract(Me, Value)
End Function

''
' Multiplies this instance of BigInteger by the supplied multiplier.
'
' @param Value A numeric representation of the value to multiply this instance by.
' @return A new instance of BigInteger containing the product value.
' @remarks This method accepts a BigInteger, one of VB's numeric values or
' a string to be parsed.
' <pre>
' Dim b As BigInteger
' Set b = BInt(100)
' Set b = b.Multiply(BInt(150))
' Debug.Print b.ToString '' 15000
' </pre>
'
Public Function Multiply(ByVal Value As BigInteger) As BigInteger
    Set Multiply = BigInteger.Multiply(Me, Value)
End Function

''
' Divides this BigInteger instance by the supplied divisor.
'
' @param Divisor A numeric representation of the value to divide this instance by.
' @return A new instance of BigInteger containing the quotient value.
' @remarks This method accepts a BigInteger, one of VB's numeric values or
' a string to be parsed.
' <pre>
' Dim b As BigInteger
' Set b = BInt(100)
' Set b = b.Divide(BInt(40))
' Debug.Print b.ToString '' 2
' </pre>
'
Public Function Divide(ByVal Divisor As BigInteger) As BigInteger
    Set Divide = BigInteger.Divide(Me, Divisor)
End Function

''
' Divides this value by the value, returning the quotient and remainding.
'
' @param Divisor A numeric representation of the value to divide this instance by.
' @param Remainder A reference to the variable to receive the remainder as a result of the division.
' @return A new instance of BigInteger containing the quotient value.
' @remarks This method accepts a BigInteger, one of VB's numeric values or
' a string to be parsed.
' <pre>
' Dim b As BigInteger
' Dim r As BigInteger
' Dim q As BigInteger
' Set b = BInt(100)
' Set q = b.DivRem(BInt(40), r)
' Debug.Print q.ToString '' 2
' Debug.Print r.ToString '' 20
' </pre>
Public Function DivRem(ByVal Divisor As BigInteger, ByRef remainder As BigInteger) As BigInteger
    Set DivRem = BigInteger.DivRem(Me, Divisor, remainder)
End Function

''
' Performs a modulus operation, returning the remainder.
'
' @param Value A numeric representation of the value to divide this instance by.
' @return The modulus (remainder) of the division.
' @remarks This method accepts a BigInteger, one of VB's numeric values or
' a string to be parsed.
' <pre>
' Dim b As BigInteger
' Set b = BInt(100)
' Set b = b.Modulus(BInt(40))
' Debug.Print b.ToString '' 20
' </pre>
'
Public Function Modulus(ByVal Value As BigInteger) As BigInteger
    Set Modulus = BigInteger.Modulus(Me, Value)
End Function

''
' Returns the greatest common divisor between this instance and the passed in value.
'
' @param Value The other value used to determine the greatest common divisor between.
' @return Returns the greatest common divisor for the two values.
'
Public Function GreatestCommonDivisor(ByVal Value As BigInteger) As BigInteger
    Set GreatestCommonDivisor = BigInteger.GreatestCommonDivisor(Me, Value)
End Function

''
' Tests a specific bit in the number to determine if it is set (1) or not (0).
'
' @param Bit The bit to be tested.
' @return Returns True if the bit is set (1), false if it is not (0).
' @remarks The max value for bit is 2147483647. If the bit value is larger than
' the BigInteger instance, then the sign of the number is used. This will represent
' a negative number with infinite sign extension bits.
'
Public Function TestBit(ByVal Bit As Long) As Boolean
    If Bit < 0 Then _
        Error.ArgumentOutOfRange "Bit", ArgumentOutOfRange_NeedNonNegNum
    
    Dim Digit As Digit
    
    Digit = MakeDigit(Bit)
    
    If Digit.Index < mNumber.Precision Then
        TestBit = mNumber.Digits(Digit.Index) And Digit.Value
    Else
        TestBit = (mNumber.Sign = Negative)
    End If
End Function

''
' Sets the specified bit of the BigInteger instance to 1.
'
' @param Bit The bit to be set to 1.
' @return A new BigInteger value with the bit set to 1.
' @remarks If the bit specified is larger than the current instance
' of the BigInteger value, then the new value will be extended to
' include the bit being set.
'
Public Function SetBit(ByVal Bit As Long) As BigInteger
    If Bit < 0 Then _
        Error.ArgumentOutOfRange "Bit", ArgumentOutOfRange_NeedNonNegNum
    
    Dim Digit   As Digit
    Dim Number  As BigNumber
    
    Digit = MakeDigit(Bit)
    Number.Sign = mNumber.Sign
    
    If Digit.Index < mNumber.Precision Then
        Number.Digits = mNumber.Digits
        Number.Precision = mNumber.Precision
    ElseIf mNumber.Sign <> -1 Then
        ReDim Number.Digits(0 To Digit.Index)
        Number.Precision = Digit.Index + 1
        Number.Sign = 1
        
        If Not IsZero Then
            CopyMemory Number.Digits(0), mNumber.Digits(0), mNumber.Precision * BytesPerDigit
        End If
    Else
        Set SetBit = Me
        Exit Function
    End If
    
    Number.Digits(Digit.Index) = Number.Digits(Digit.Index) Or Digit.Value
    
    Set SetBit = New BigInteger
    SetBit.InitBigNumber Number
End Function

''
' Clears the specified bit of the BigInteger instance to 0.
'
' @param Bit The bit to be set to 0.
' @return A new BigInteger value with the bit set to 0.
' @remarks Bit is a zero-based index. The first bit is 0.
Public Function ClearBit(ByVal Bit As Long) As BigInteger
    If Bit < 0 Then _
        Error.ArgumentOutOfRange "Bit", ArgumentOutOfRange_NeedNonNegNum
    
    Dim Digit   As Digit
    Dim Number  As BigNumber

    If IsZero Then
        Set ClearBit = BigInteger.Zero
        Exit Function
    End If
    
    Digit = MakeDigit(Bit)
    Number.Sign = mNumber.Sign
    
    If Digit.Index < mNumber.Precision Then
        Number.Digits = mNumber.Digits
        Number.Precision = mNumber.Precision
    ElseIf Sign = -1 Then
        ReDim Number.Digits(0 To Digit.Index)
        CopyMemory Number.Digits(0), mNumber.Digits(0), mNumber.Precision * BytesPerDigit
        FillMemory Number.Digits(mNumber.Precision), (Digit.Index - mNumber.Precision + 1) * BytesPerDigit, &HFF
        Number.Precision = Digit.Index + 1
    Else
        Set ClearBit = Me
        Exit Function
    End If
    
    Number.Digits(Digit.Index) = Number.Digits(Digit.Index) And (Not Digit.Value)
    
    Set ClearBit = CreateNormalizedBigInteger(Number)
End Function

Private Function MakeDigit(ByVal Bit As Long) As Digit
    MakeDigit.Index = (Bit And &H7FFFFFF0) \ vbShift4Bits
    MakeDigit.Value = PowersOf2(Bit And &HF)
End Function

Private Function CreateNormalizedBigInteger(ByRef Number As BigNumber) As BigInteger
    Number.Precision = FindPrecision(Number)
    
    If Number.Precision = 0 Then
        Set CreateNormalizedBigInteger = BigInteger.Zero
    Else
        Set CreateNormalizedBigInteger = New BigInteger
        CreateNormalizedBigInteger.InitBigNumber Number
    End If
End Function

Private Function FindPrecision(ByRef Number As BigNumber) As Long
    Dim i As Long
    
    For i = Number.Precision - 1 To 0 Step -1
        If Number.Digits(i) <> 0 Then
            FindPrecision = i + 1
            Exit Function
        End If
    Next
End Function

''
' Flips the specified bit, toggling between True and False.
'
' @param Bit The bit to toggle.
' @return A new BigInteger with the specified bit flipped.
' @remarks If the bit specified is larger than the current instance
' of the BigInteger value, then the new value will be extended to
' include the bit being flipped.
'
Public Function FlipBit(ByVal Bit As Long) As BigInteger
    If TestBit(Bit) Then
        Set FlipBit = ClearBit(Bit)
    Else
        Set FlipBit = SetBit(Bit)
    End If
End Function

''
' Peforms a bitwise AND of this BigInteger instance and the value passed in.
'
' @param Value The value to perform the bitwise AND with.
' @return A new BigInteger value containing the performed bitwise AND result.
'
Public Function AndBits(ByVal Value As BigInteger) As BigInteger
    Set AndBits = BigInteger.AndBits(Me, Value)
End Function

''
' Performs a bitwise OR of this BigInteger instance and the value passed in.
'
' @param Value The value to perform the bitwise OR with.
' @return A new BigInteger value containing the performed bitwise OR result.
'
Public Function OrBits(ByVal Value As BigInteger) As BigInteger
    Set OrBits = BigInteger.OrBits(Me, Value)
End Function

''
' Performs a bitwise XOR of this BigInteger instance and the value passed in.
'
' @param Value The value to perform the bitwise XOR with.
' @return A new BigInteger value containing the performed bitwise XOR result.
'
Public Function XorBits(ByVal Value As BigInteger) As BigInteger
    Set XorBits = BigInteger.XorBits(Me, Value)
End Function

''
' Peforms a bitwise NOT of this BigInteger instance and the value passed in.
'
' @param Value The value to perform the bitwise NOT with.
' @return A new BigInteger value containing the performed bitwise NOT result.
'
Public Function NotBits() As BigInteger
    If IsZero Then
        Set NotBits = BigInteger.MinusOne
        Exit Function
    ElseIf IsMinusOne Then
        Set NotBits = BigInteger.Zero
        Exit Function
    End If
    
    Set NotBits = New BigInteger
    NotBits.InitNotBits mNumber
End Function

''
' Peforms a bitwise AND of this BigInteger instance and the NOT value passed in.
'
' @param Value The value that will be NOTted to perform the bitwise AND with.
' @return A new BigInteger value containing the performed bitwise AND/NOT result.
'
Public Function AndNotBits(ByVal Value As BigInteger) As BigInteger
    Set AndNotBits = BigInteger.AndNotBits(Me, Value)
End Function

''
' Raises the BigInteger value to the specified power.
'
' @param Exponent The value to raise the BigInteger value to.
' @return Returns the BigInteger value raised to the specified power.
'
Public Function Pow(ByVal Exponent As BigInteger) As BigInteger
    Set Pow = BigInteger.Pow(Me, Exponent)
End Function

''
' Performs a bit shifting operation to the the right.
'
' @param BitCount The number of bits the BigInteger value is to be shifted.
' @return Returns a new BigInteger instance containing the bit shifted value.
' @remarks If BitCount is negative, then a left shift is performed.
Public Function ShiftRight(ByVal BitCount As Long) As BigInteger
    If (BitCount = 0) Or IsZero Then
        Set ShiftRight = Me
        Exit Function
    ElseIf BitCount < 0 Then
        Set ShiftRight = ShiftLeft(-BitCount)
        Exit Function
    End If
    
    Dim WordCount As Long
    WordCount = (BitCount And &HFFFFFFF0) \ vbShift4Bits

    If WordCount >= mNumber.Precision Then
        If mNumber.Sign = Negative Then
            Set ShiftRight = BigInteger.MinusOne
        Else
            Set ShiftRight = BigInteger.Zero
        End If
        Exit Function
    End If
    
    Set ShiftRight = New BigInteger
    ShiftRight.InitShiftRight mNumber, BitCount
End Function

''
' Performs a left shift operation on the BigInteger instance.
'
' @param BitCount The number of bits the value is to be shifted left.
' @return Returns a new BigInteger instance containing the shifted value.
' @remarks If BitCount is negative, then a right shift is performed.
Public Function ShiftLeft(ByVal BitCount As Long) As BigInteger
    If (BitCount = 0) Or IsZero Then
        Set ShiftLeft = Me
        Exit Function
    End If
    
    Set ShiftLeft = New BigInteger
    
    If BitCount > 0 Then
        ShiftLeft.InitShiftLeft mNumber, BitCount
    Else
        ShiftLeft.InitShiftRight mNumber, -BitCount
    End If
End Function

''
' Returns the integer square root of a BigInteger.
'
' @Return The integer square root of the value.
'
Public Function SquareRoot() As BigInteger
    Set SquareRoot = BigInteger.SquareRoot(Me)
End Function

''
' Returns the mininum of this instance and a second BigInteger instance.
'
' @param value One of the two values to be tested.
' @return If this instance is less than or equal to <i>value</i>,
' then this instance is returned, otherwise <i>value</i> is returned.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function Min(ByVal Value As BigInteger) As BigInteger
    Set Min = BigInteger.Min(Me, Value)
End Function

''
' Returns the maximum of this instance and a second BigInteger instance.
'
' @param value One of the two values to be tested.
' @return If this instance is equal to or greater than <i>value</i>,
' then this instance is returned, otherwise <i>value</i> is returned.
' <p>Values of Nothing will be treated as zero.</p>
'
Public Function Max(ByVal Value As BigInteger) As BigInteger
    Set Max = BigInteger.Max(Me, Value)
End Function

''
' Returns the factorial of this BigInteger value.
'
' @return The factorial of this BigInteger instance value.
'
Public Function Factorial() As BigInteger
    Set Factorial = BigInteger.Factorial(Me)
End Function


Friend Sub AttachNumber(ByRef n As BigNumber)
    CopyMemory n, mNumber, LenB(n)
End Sub

Friend Sub DetachNumber(ByRef n As BigNumber)
    ZeroMemory n, LenB(n)
End Sub

Friend Sub CopyNumber(ByRef n As BigNumber)
    n = mNumber
End Sub

Friend Property Get IsMinusOne() As Boolean
    If mNumber.Sign = -1 And mNumber.Precision = 1 Then
        IsMinusOne = mNumber.Digits(0) = -1
    End If
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            InitFromLong Value
        Case vbDouble, vbSingle
            InitFromFloat Value
        Case vbCurrency
            InitFromCurrency Value
        Case vbDecimal
            InitFromDecimal Value
        Case vbByte Or vbArray
            If CorArray.IsNull(Value) Then _
                Error.ArgumentNull "Value"
        
            InitFromArray Value
        Case Else
            Throw Cor.NewArgumentException("Cannot initialize BigInteger from specified value.")
    End Select
    
    If mNumber.Sign <> 0 Then
        mNumber.Precision = UBound(mNumber.Digits) + 1
    End If
End Sub

Friend Sub InitBigNumber(ByRef n As BigNumber)
    CopyMemory mNumber, n, LenB(n)
    ZeroMemory n, LenB(n)
End Sub

Friend Sub InitNotBits(ByRef Other As BigNumber)
    ReDim mNumber.Digits(0 To Other.Precision) ' keep one extra digit incase we need to add leading negative FFFF.
    mNumber.Precision = Other.Precision
    mNumber.Sign = -Other.Sign
    
    Dim i As Long
    For i = 0 To mNumber.Precision - 1
        mNumber.Digits(i) = Not Other.Digits(i)
    Next i
    
    If mNumber.Sign = -1 Then
        If (mNumber.Digits(mNumber.Precision - 1) And &H8000) = 0 Then
            mNumber.Digits(mNumber.Precision) = &HFFFF
            mNumber.Precision = mNumber.Precision + 1
        End If
    Else
        TrimLeadingZeros
    End If
End Sub

Friend Sub InitShiftLeft(ByRef Other As BigNumber, ByVal ShiftCount As Long)
    Debug.Assert ShiftCount > 0 And Other.Sign <> 0
    mNumber.Sign = Other.Sign
    mNumber.Digits = ShiftLeftNumber(Other.Digits, ShiftCount)
    mNumber.Precision = UBound(mNumber.Digits) + 1
    TrimLeadingPadding
End Sub

Friend Sub InitShiftRight(ByRef Other As BigNumber, ByVal ShiftCount As Long)
    Debug.Assert ShiftCount > 0 And Other.Sign <> 0
    mNumber.Sign = Other.Sign
    mNumber.Digits = ShiftRightNumber(Other.Digits, ShiftCount)
    mNumber.Precision = UBound(mNumber.Digits) + 1
    TrimLeadingPadding
End Sub

Friend Sub InternalInit(ByRef Bytes() As Integer)
    SAPtr(mNumber.Digits) = SAPtr(Bytes)
    SAPtr(Bytes) = vbNullPtr
    Normalize
End Sub

Friend Sub InternalInitNumber(ByRef n As BigNumber)
    CopyMemory mNumber, n, LenB(n)
    ZeroMemory n, LenB(n)
End Sub

Friend Sub InternalInitNumber2(ByRef n As BigNumber)
    CopyMemory mNumber, n, LenB(n)
    ZeroMemory n, LenB(n)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValueOrZero(ByVal Value As BigInteger) As BigInteger
    If Value Is Nothing Then
        Set ValueOrZero = BigInteger.Zero
    Else
        Set ValueOrZero = Value
    End If
End Function

Private Sub InitFromLong(ByVal Value As Long)
    Dim IsSingleDigit As Boolean
    
    If Value < 0 Then
        mNumber.Sign = -1
        IsSingleDigit = (Value And &HFFFF8000) = &H8000
    ElseIf Value > 0 Then
        mNumber.Sign = 1
        IsSingleDigit = (Value And &HFFFF0000) = 0
    Else
        Exit Sub
    End If
    
    If IsSingleDigit Then
        ReDim mNumber.Digits(0 To 0)
        mNumber.Digits(0) = AsWord(Value)
    Else
        ReDim mNumber.Digits(0 To 1)
        AsLong(mNumber.Digits(0)) = Value
    End If
End Sub

Private Sub InitFromArray(ByRef Value As Variant)
    Dim Bytes() As Byte
    
    If Len1D(Value) = 0 Then
        Exit Sub
    End If

    SAPtr(Bytes) = SAPtrV(Value)
    InitFromBytes Bytes
    SAPtr(Bytes) = vbNullPtr
End Sub

Private Sub InitFromBytes(ByRef Bytes() As Byte)
    Dim ByteCount   As Long
    Dim MaxIndex    As Long
    Dim MinIndex    As Long
    Dim Size        As Long
    
    MaxIndex = UBound(Bytes)
    MinIndex = LBound(Bytes)
    
    Select Case Bytes(MaxIndex)
        Case 0
            ' remove all zeros; sign is positive or zero
            Do While MaxIndex >= MinIndex
                If Bytes(MaxIndex) > 0 Then
                    Exit Do
                End If
                
                MaxIndex = MaxIndex - 1
            Loop
            
            If MaxIndex >= MinIndex Then
                mNumber.Sign = 1
            End If
        Case &HFF
            ' remove all leading FF unless first non FF does not have high-bit set.
            ' FFFFFF => FF
            ' FFFF80 => 80
            ' FFFF01 => FF01
            Do While MaxIndex > MinIndex And Bytes(MaxIndex) = &HFF
                If (Bytes(MaxIndex - 1) And &H80) = 0 Then
                    Exit Do
                End If
                
                MaxIndex = MaxIndex - 1
            Loop
            
            mNumber.Sign = -1
        Case &H80
            ' leave; sign is negative
            mNumber.Sign = -1
        Case Else
            mNumber.Sign = 1
    End Select
    
    ByteCount = MaxIndex - MinIndex + 1
    Size = CorMath.Ceiling(ByteCount / 2)
    
    If Size > 0 Then
        ReDim mNumber.Digits(0 To Size - 1)
        CopyMemory mNumber.Digits(0), Bytes(MinIndex), ByteCount
        
        If mNumber.Sign = -1 Then
            If (mNumber.Digits(Size - 1) And &HFF80) = &H80 Then
                mNumber.Digits(Size - 1) = mNumber.Digits(Size - 1) Or &HFF00
            End If
        End If
    End If
End Sub

Private Sub InitFromFloat(ByVal Value As Double)
    Dim Exponent    As Long
    Dim InputDigits As QWordArray
        
    Value = Fix(Value)
    InputDigits = AsQWordArr(Value)
    mNumber.Sign = Sgn(Int(Value))
    Exponent = (InputDigits.Words(3) And &H7FF0) \ &H10
    
    If Exponent = 0 Then
        Exit Sub
    End If

    ' remove the sign bit and exponent values, and
    ' add the leading 1 back to the number.
    InputDigits.Words(3) = (InputDigits.Words(3) And &HF) Or &H10

    If Exponent > &H433 Then
        mNumber.Digits = ShiftLeftNumber(InputDigits.Words, Exponent - &H433)
    Else
        mNumber.Digits = ShiftRightNumber(InputDigits.Words, &H433 - Exponent)
    End If

    If mNumber.Sign < 0 Then
        ApplyTwosComplement mNumber.Digits
    End If
    
    TrimLeadingPadding
End Sub

Private Sub InitFromCurrency(ByVal Value As Currency)
    Value = Fix(Value) * 0.0001@
    ReDim mNumber.Digits(0 To 3)
    AsCurr(mNumber.Digits(0)) = Value
    mNumber.Sign = Sgn(Value)
    TrimLeadingPadding
End Sub

' VB Decimal layout
' bytes 0-1: Variant datatype information
' bytes 2: precision
' bytes 3: sign (&h80 is negative)
' bytes 4-7: the 32 highest bits in the 96bit value
' bytes 8-11: the 32 lowest bits in the 96bit value
' bytes 12-15: the 32 middle bits in the 96bit value
'
' .NET Decimal layout
' bytes 0-3: the 32 lowest bits in the 96bit value
' bytes 4-7: the 32 middle bits in the 96bit value
' bytes 8-11: the 32 highest bits in the 96bit value
' bytes 12-13: unused (zero)
' bytes 14: precision
' bytes 15: sign (&h80 is negative)
Private Sub InitFromDecimal(ByRef Value As Variant)
    Dim Ptr As Long
    
    Value = Fix(Value)
    Ptr = VarPtr(Value)
    ReDim mNumber.Digits(0 To 5)
    AsLong(mNumber.Digits(0)) = MemLong(Ptr + 8)
    AsLong(mNumber.Digits(2)) = MemLong(Ptr + 12)
    AsLong(mNumber.Digits(4)) = MemLong(Ptr + 4)
    mNumber.Sign = Sgn(Value)
    
    If mNumber.Sign < 0 Then
        ApplyTwosComplement mNumber.Digits
    End If
    
    TrimLeadingPadding
End Sub

Private Sub TrimLeadingPadding()
    If mNumber.Sign < 0 Then
        TrimNegativeDigits
    Else
        TrimLeadingZeros
    End If
End Sub

Private Sub TrimNegativeDigits()
    Dim Index As Long
            
    Index = UBound(mNumber.Digits)
    ' remove leading FFFF unless first non FFFF does not have high bit set.
    ' FFFFFFFFFFFF => FFFF
    ' FFFFFFFF8000 => 8000
    ' FFFFFFFF0001 => FFFF0001
    Do While Index > 0 And mNumber.Digits(Index) = &HFFFF
        If (mNumber.Digits(Index - 1) And &H8000) = 0 Then
            Exit Do
        End If
        
        Index = Index - 1
    Loop
            
    ReDim Preserve mNumber.Digits(0 To Index)
    mNumber.Precision = Index + 1
End Sub

Private Sub TrimLeadingZeros()
    Dim StartIndex  As Long
    Dim Index       As Long
    
    StartIndex = UBound(mNumber.Digits)
    Index = StartIndex
    
    Do While Index >= 0
        If mNumber.Digits(Index) <> 0 Then
            Exit Do
        End If
        
        Index = Index - 1
    Loop
    
    If Index < 0 Then
        Erase mNumber.Digits
        mNumber.Sign = 0
        mNumber.Precision = 0
    ElseIf Index < StartIndex Then
        ReDim Preserve mNumber.Digits(0 To Index)
        mNumber.Precision = Index + 1
    End If
End Sub

Private Sub Normalize(Optional ByVal ub As Long = -1)
    If ub < 0 Then ub = UBound(mNumber.Digits)
    If ub < 0 Then Exit Sub

    Dim i As Long
    Select Case mNumber.Digits(ub)
        Case 0   ' we have a leading zero digit

            ' now search for the first nonzero digit from the left.
            For i = ub - 1 To 0 Step -1
                If mNumber.Digits(i) <> 0 Then
                    ' we found a nonzero digit, so set the number
                    mNumber.Sign = Positive     ' we know it's positive because of the leading zero
                    mNumber.Precision = i + 1   ' set the number of digits
                    Exit Sub
                End If
            Next i

            mNumber.Sign = Zero
            mNumber.Precision = 0

        Case &HFFFF ' we have a leading negative

            mNumber.Sign = Negative ' we know this for sure

            For i = ub To 0 Step -1
                If mNumber.Digits(i) <> &HFFFF Then
                    If mNumber.Digits(i) And &H8000 Then
                        mNumber.Precision = i + 1
                    Else
                        mNumber.Precision = i + 2
                    End If
                    Exit Sub
                End If
            Next i

            ' the array was full of &HFFFF, we only need to represent one.
            mNumber.Precision = 1

        Case Else
            If mNumber.Digits(ub) And &H8000 Then
                mNumber.Sign = Negative
            Else
                mNumber.Sign = Positive
            End If

            mNumber.Precision = ub + 1

    End Select
End Sub

Private Function ToDecimalString(ByVal Count As Long, ByVal FormatInfo As NumberFormatInfo) As String
    Const DIGITS_PER_INTEGER As Long = 5
    Dim Dividend    As BigNumber
    Dim Size        As Long
    Dim Chars()     As Integer
    Dim Output      As String
    Dim pos         As Long
    
    Dividend = mNumber
    
    If Dividend.Sign = -1 Then
        BigNumberMath.ApplyTwosComplement Dividend.Digits
    End If
        
    Size = Dividend.Precision * DIGITS_PER_INTEGER
    
    If Count > Size Then
        Size = Count
    End If
    
    Output = String$(Size + 1, vbZeroChar)
    Chars = AsChars(Output)
    
    pos = Len(Output) - 1
    
    Do While Dividend.Precision > 0
        Chars(pos) = SingleInPlaceDivideBy10(Dividend) + vbZeroChar
        pos = pos - 1
    Loop
    
    If Size - pos < Count Then
        pos = Size - Count
    End If
    
    If mNumber.Sign = Negative Then
        Chars(pos) = AscW(FormatInfo.NegativeSign)
        pos = pos - 1
    End If
    
    FreeChars Chars
    ToDecimalString = Mid$(Output, pos + 2)
End Function

Private Function ToHexString(ByVal Count As Long, Offset As Long) As String
    Const DIGITS_PER_INTEGER As Long = 4
    
    Dim Size    As Long
    Dim Output  As String
    
    Size = mNumber.Precision * DIGITS_PER_INTEGER
    
    If Count > Size Then
        Size = Count
    End If
    
    Output = String$(Size, vbZeroChar)
    
    Dim c As Long
    Dim k As Long
    Dim i As Long
    For i = 0 To mNumber.Precision - 1
        k = mNumber.Digits(i) And &HFFFF&
        c = Size - i * DIGITS_PER_INTEGER
        
        Do While k > 0
            If k And &HF Then
                Mid$(Output, c, 1) = Mid$(HEX_CHARS, (k And &HF) + Offset, 1)
            End If
            
            k = k \ &H10
            c = c - 1
        Loop
    Next
    
    If Size - c < Count Then
        If mNumber.Sign = Negative Then
            Mid$(Output, Count - Size + 1) = String$(Count - (Size - c), Mid$(HEX_CHARS, &HF + Offset, 1))
        End If
        
        c = Size - Count
    End If
    
    ToHexString = Mid$(Output, c + 1)
End Function

Private Function ToBinaryString(ByVal Count As Long) As String
    Const DIGITS_PER_INTEGER As Long = 16
    
    Dim Size    As Long
    Dim Output  As String
    
    Size = mNumber.Precision * DIGITS_PER_INTEGER + 1
    
    If Count > Size Then
        Size = Count
    End If
    
    Output = String$(Size, vbZeroChar)
    
    Dim c As Long
    Dim k As Long
    Dim i As Long
    For i = 0 To mNumber.Precision - 1
        k = mNumber.Digits(i) And &HFFFF&
        c = Size - i * DIGITS_PER_INTEGER
         
        Do While k > 0
            If k And 1 Then Mid$(Output, c, 1) = "1"
            k = k \ &H2
            c = c - 1
        Loop
    Next i
    
    If Size - c < Count Then
        If mNumber.Sign = Negative Then
            Mid$(Output, Count - Size + 1) = String$(Count - (Size - c), vbOneChar)
        End If
        
        c = Size - Count
    End If
    
    ToBinaryString = Mid$(Output, c + 1)
End Function

Private Function ParseFormatSpecifier(ByRef Format As String, ByRef OutSpecifier As String) As Long
    Dim Chars() As Integer
    
    If Len(Format) = 0 Then
        OutSpecifier = "d"
        ParseFormatSpecifier = 1
        Exit Function
    End If
    
    On Error GoTo Catch
    Chars = AllocChars(Format)
    
    Select Case Chars(0)
        Case vbLowerXChar, vbUpperXChar, vbLowerBChar, vbUpperBChar, vbLowerDChar, vbUpperDChar
            OutSpecifier = Chr$(Chars(0))
        Case Else
            Error.Format Format_InvalidString
    End Select
            
    Dim Result  As Long
    Dim i       As Long
    For i = 1 To UBound(Chars)
        Dim Ch As Long
        Ch = Chars(i)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                Result = Result * 10 + (Ch - vbZeroChar)
                
                If Result > 1000000 Then
                    Exit For   ' hopefully one million digits is enough.
                End If
            Case Else
                Error.Format Format_InvalidString
        End Select
    Next i
    
    ParseFormatSpecifier = IIfLong(Result = 0, 1, Result)
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
    
Finally:
    FreeChars Chars
    Return
End Function

Private Function ShiftLeftNumber(ByRef Digits() As Integer, ByVal BitCount As Long) As Integer()
    Dim WholeWords      As Long
    Dim RemainingBits   As Long
    Dim Size            As Long
    Dim NewSize         As Long
    Dim NewNumber()     As Integer
    
    Size = Len1D(Digits)
    WholeWords = (BitCount And &HFFFFFFF0) \ &H10
    RemainingBits = BitCount And &HF
    NewSize = Size + WholeWords + IIfLong(RemainingBits > 0, 1, 0)
    ReDim NewNumber(0 To NewSize - 1)
    
    If RemainingBits = 0 Then
        CopyMemory NewNumber(WholeWords), Digits(0), Size * 2
    Else
        Dim LowerMask   As Long
        Dim UpperMask   As Long
        Dim RShift      As Long
        Dim LShift      As Long
        
        LowerMask = ((PowersOf2(16 - RemainingBits) And &HFFFF&) - 1) And &HFFFF&
        UpperMask = (Not LowerMask) And &HFFFF&
        RShift = PowersOf2(16 - RemainingBits) And &HFFFF&
        LShift = PowersOf2(RemainingBits) And &HFFFF&
           
        Dim SrcInt      As Long
        Dim DstInt      As Long
        Dim i           As Long
        Dim j           As Long

        i = WholeWords
        Do
            SrcInt = Digits(j) And &HFFFF&
            DstInt = DstInt Or ((SrcInt And LowerMask) * LShift)
            
            #If Release Then
            NewNumber(i) = DstInt
            #Else
            NewNumber(i) = AsWord(DstInt)
            #End If
            
            DstInt = (SrcInt And UpperMask) \ RShift
            i = i + 1
            j = j + 1
        Loop While j < Size
        
        NewNumber(i) = DstInt
    End If

    ShiftLeftNumber = NewNumber
End Function

Private Function ShiftRightNumber(ByRef Digits() As Integer, ByVal BitCount As Long) As Integer()
    Dim WordCount       As Long
    Dim RemainingBits   As Long
    Dim NewPrecision    As Long
    Dim NewNumber()     As Integer
    Dim Size            As Long
    
    Size = Len1D(Digits)
    WordCount = (BitCount And &HFFFFFFF0) \ &H10
    RemainingBits = BitCount And &HF
    NewPrecision = Size - WordCount
    ReDim NewNumber(0 To NewPrecision - 1)
    
    If RemainingBits = 0 Then
        CopyMemory NewNumber(0), Digits(WordCount), NewPrecision * 2
    Else
        Dim LowerMask   As Long
        Dim UpperMask   As Long
        Dim RShift      As Long
        Dim LShift      As Long
        
        LowerMask = ((PowersOf2(RemainingBits) And &HFFFF&) - 1) And &HFFFF&
        UpperMask = (Not LowerMask) And &HFFFF&
        RShift = PowersOf2(RemainingBits) And &HFFFF&
        LShift = PowersOf2(16 - RemainingBits) And &HFFFF&
            
        Dim SrcInt      As Integer
        Dim DstInt      As Integer
        Dim i           As Long
        Dim j           As Long
        
        j = WordCount
        SrcInt = Digits(j)
        Do
            DstInt = ((SrcInt And &HFFFF&) And UpperMask) \ RShift
            j = j + 1
            
            If j = Size Then
                Exit Do
            End If
            
            SrcInt = Digits(j)
            NewNumber(i) = DstInt Or AsWord(((SrcInt And &HFFFF&) And LowerMask) * LShift)
            i = i + 1
        Loop
            
        NewNumber(i) = DstInt
    End If

    ShiftRightNumber = NewNumber
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Cor.BInt(Value))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(Format As String, ByVal Provider As IFormatProvider) As String
    IFormattable_ToString = ToString(Format, Provider)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
