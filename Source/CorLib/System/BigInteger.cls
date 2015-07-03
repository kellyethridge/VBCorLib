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
Implements IComparable

Private Const HEX_CHARS As String = "0123456789abcdef0123456789ABCDEF"
Private Const LOWER_HEX As Long = 1
Private Const UPPER_HEX As Long = 17

Private mNumber As modBigMath.Number


''
' Returns the sign of the number this instance represents.
'
' @return The value of type Sign representing the sign of the number.
' @remarks Returns -1 for negative, 0 for zero, 1 for positive.
'
Public Property Get Sign() As Sign
    Sign = mNumber.Sign
End Property

''
' Returns the number of bytes being used to represent the number.
'
' @return The number of bytes being used to represent the number.
' @remarks There may be additional bytes allocated, but not considered
' as part of the number. Those bytes are not represented in the precision.
' <p>Since a negative value can have an infinite number of leading &HFFFF values,
' only 1 (2 bytes) is counted if it is needed.</p>
'
Public Property Get Precision() As Long
    Dim Ret As Long
    
    Ret = mNumber.Precision * 2
    If Ret > 0 Then
        If (mNumber.Digits(mNumber.Precision - 1) And &HFF00) = 0 Then
            Ret = Ret - 1
        End If
    End If
    
    Precision = Ret
End Property

''
' Returns a byte array representing the number.
'
' @return A byte array the same length as Precision that represents the number.
' @remarks To represent a byte array for a negative value, the highest bit must be
' set. To ensure this, a pair of bytes with the value of &HFF may be part of the return
' value.
'
Public Function ToArray() As Byte()
    Dim Ret() As Byte
    
    If mNumber.Precision = 0 Then
        Ret = Cor.NewBytes()
    Else
        Dim i As Long
        i = Me.Precision
        
        ReDim Ret(0 To i - 1)
        Call Buffer.BlockCopy(mNumber.Digits, 0, Ret, 0, i)
    End If
        
    ToArray = Ret
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
Public Function ToString(Optional ByVal Format As String) As String
    Dim Specifier   As String
    Dim Count       As Long
    Count = ParseFormatSpecifier(Format, Specifier)
    If Count = 0 Then Count = 1
    
    Select Case Specifier
        Case "x":       ToString = ToHexString(Count, LOWER_HEX)
        Case "X":       ToString = ToHexString(Count, UPPER_HEX)
        Case "b", "B":  ToString = ToBinaryString(Count)
        Case "d", "D":  ToString = ToDecimalString(Count)
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
Public Function Equals(ByVal Other As BigInteger) As Boolean
    Equals = BigInteger.Equals(Me, Other)
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
        Dim num As Number
        num = mNumber
        
        Call modBigMath.Negate(num)
        
        Set Negate = New BigInteger
        Call Negate.InternalInitNumber(num)
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
        Set Absolute = Me.Negate
    Else
        Set Absolute = Me
    End If
End Function

''
' Adds a value to this instance of BigInteger.
'
' @param value A numeric representation of the value to be added to this instance.
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
' @param value A numeric representation of the value to be subtracted to this instance.
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
' @param value A numeric representation of the value to multiply this instance by.
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
' @param value A numeric representation of the value to divide this instance by.
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
Public Function Divide(ByVal Value As BigInteger) As BigInteger
    Set Divide = BigInteger.Divide(Me, Value)
End Function

''
' Divides this value by the value, returning the quotient and remainding.
'
' @param value A numeric representation of the value to divide this instance by.
' @param remainder A reference to the variable to receive the remainder as a result of the division.
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
'
Public Function DivRem(ByVal Value As BigInteger, ByRef remainder As BigInteger) As BigInteger
    Set DivRem = BigInteger.DivRem(Me, Value, remainder)
End Function

''
' Performs a modulus operation, returning the remainder.
'
' @param value A numeric representation of the value to divide this instance by.
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
' @param value The other value used to determine the greatest common divisor between.
' @return Returns the greatest common divisor for the two values.
'
Public Function GreatestCommonDivisor(ByVal Value As BigInteger) As BigInteger
    Set GreatestCommonDivisor = BigInteger.GreatestCommonDivisor(Me, Value)
End Function

''
' Tests a specific bit in the number to determine if it is set (1) or not (0).
'
' @param bit The bit to be tested.
' @return Returns True if the bit is set (1), false if it is not (0).
' @remarks The max value for bit is 2147483647. If the bit value is larger than
' the BigInteger instance, then the sign of the number is used. This will represent
' a negative number with infinite sign extension bits.
'
Public Function TestBit(ByVal bit As Long) As Boolean
    If bit < 0 Then _
        Call Err.Raise(5, , "Bit cannot be negative.")
    
    If mNumber.Precision = 0 Then Exit Function
    
    Dim Index As Long
    Index = (bit And &HFFFFFFF0) \ &H10
    
    If Index < mNumber.Precision Then
        TestBit = mNumber.Digits(Index) And (PowersOf2(bit And &HF))
    Else
        TestBit = (mNumber.Sign = Negative)
    End If
End Function

''
' Sets the specified bit of the BigInteger instance to 1.
'
' @param bit The bit to be set to 1.
' @return A new BigInteger value with the bit set to 1.
' @remarks If the bit specified is larger than the current instance
' of the BigInteger value, then the new value will be extended to
' include the bit being set.
'
Public Function SetBit(ByVal bit As Long) As BigInteger
    Set SetBit = InternalSetBit(bit, True)
End Function

''
' Clears the specified bit of the BigInteger instance to 0.
'
' @param bit The bit to be set to 0.
' @return A new BigInteger value with the bit set to 0.
'
Public Function ClearBit(ByVal bit As Long) As BigInteger
    Set ClearBit = InternalSetBit(bit, False)
End Function

''
' Flips the specified bit, toggling between True and False.
'
' @param bit The bit to toggle.
' @return A new BigInteger with the specified bit flipped.
' @remarks If the bit specified is larger than the current instance
' of the BigInteger value, then the new value will be extended to
' include the bit being flipped.
'
Public Function FlipBit(ByVal bit As Long) As BigInteger
    Set FlipBit = InternalSetBit(bit, Not TestBit(bit))
End Function

''
' Peforms a bitwise AND of this BigInteger instance and the value passed in.
'
' @param value The value to perform the bitwise AND with.
' @return A new BigInteger value containing the performed bitwise AND result.
'
Public Function AndBits(ByVal Value As BigInteger) As BigInteger
    Set AndBits = BigInteger.AndBits(Me, Value)
End Function

''
' Performs a bitwise OR of this BigInteger instance and the value passed in.
'
' @param value The value to perform the bitwise OR with.
' @return A new BigInteger value containing the performed bitwise OR result.
'
Public Function OrBits(ByVal Value As BigInteger) As BigInteger
    Set OrBits = BigInteger.OrBits(Me, Value)
End Function

''
' Performs a bitwise XOR of this BigInteger instance and the value passed in.
'
' @param value The value to perform the bitwise XOR with.
' @return A new BigInteger value containing the performed bitwise XOR result.
'
Public Function XorBits(ByVal Value As BigInteger) As BigInteger
    Set XorBits = BigInteger.XorBits(Me, Value)
End Function

''
' Peforms a bitwise NOT of this BigInteger instance and the value passed in.
'
' @param value The value to perform the bitwise NOT with.
' @return A new BigInteger value containing the performed bitwise NOT result.
'
Public Function NotBits() As BigInteger
    Set NotBits = BigInteger.NotBits(Me)
End Function

''
' Peforms a bitwise AND of this BigInteger instance and the NOT value passed in.
'
' @param value The value that will be NOTted to perform the bitwise AND with.
' @return A new BigInteger value containing the performed bitwise AND/NOT result.
'
Public Function AndNotBits(ByVal Value As BigInteger) As BigInteger
    Set AndNotBits = BigInteger.AndNotBits(Me, Value)
End Function

''
' Raises the BigInteger value to the specified power.
'
' @param exponent The value to raise the BigInteger value to.
' @return Returns the BigInteger value raised to the specified power.
'
Public Function Pow(ByVal exponent As BigInteger) As BigInteger
    Set Pow = BigInteger.Pow(Me, exponent)
End Function

''
' Performs a bit shifting operation to the the right.
'
' @param bitCount The number of bits the BigInteger value is to be shifted.
' @return Returns a new BigInteger instance containing the bit shifted value.
' @remarks If bitCount is negative, then a left shift is performed.
'
Public Function ShiftRight(ByVal bitCount As Long) As BigInteger
    If (bitCount = 0) Or (mNumber.Sign = Zero) Then
        Set ShiftRight = Me
        Exit Function
    ElseIf bitCount < 0 Then
        Set ShiftRight = ShiftLeft(-bitCount)
        Exit Function
    End If
    
    Dim wordCount As Long
    wordCount = (bitCount And &HFFFFFFF0) \ &H10

    If wordCount >= mNumber.Precision Then
        If mNumber.Sign = Negative Then
            Set ShiftRight = BigInteger.MinusOne
        Else
            Set ShiftRight = BigInteger.Zero
        End If
        Exit Function
    End If
    
    Set ShiftRight = New BigInteger
    Call ShiftRight.InternalInit(ShiftRightNumber(mNumber, bitCount))
End Function

''
' Performs a left shift operation on the BigInteger instance.
'
' @param bitCount The number of bits the value is to be shifted left.
' @return Returns a new BigInteger instance containing the shifted value.
' @remarks If bitCount is negative, then a right shift is performed.
'
Public Function ShiftLeft(ByVal bitCount As Long) As BigInteger
    If (bitCount = 0) Or (mNumber.Sign = Zero) Then
        Set ShiftLeft = Me
        Exit Function
    ElseIf bitCount < 0 Then
        Set ShiftLeft = ShiftRight(-bitCount)
        Exit Function
    End If
    
    Set ShiftLeft = New BigInteger
    Call ShiftLeft.InternalInit(ShiftLeftNumber(mNumber, bitCount))
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte: Call InitFromNumber(Value)
        Case vbDouble, vbSingle:        Call InitFromFloat(Value)
        Case vbCurrency:                Call InitFromCurrency(Value)
        Case vbDecimal:                 Call InitFromDecimal(Value)
        Case vbByte Or vbArray:         Call InitFromArray(Value)
        Case Else
            Throw Cor.NewArgumentException("Cannot initialize BigInteger from specified value.")
    End Select
End Sub

Friend Sub InternalInit(ByRef b() As Integer)
    SAPtr(mNumber.Digits) = SAPtr(b)
    SAPtr(b) = vbNullPtr
    Call Normalize
End Sub

Friend Sub InternalInitNumber(ByRef n As Number)
    Call CopyMemory(mNumber, n, LenB(n))
    Call ZeroMemory(n, LenB(n))
    Call Normalize(mNumber.Precision - 1)
End Sub

Friend Sub AttachNumber(ByRef n As Number)
    Call CopyMemory(n, mNumber, LenB(n))
End Sub

Friend Sub DetachNumber(ByRef n As Number)
    Call ZeroMemory(n, LenB(n))
End Sub

Friend Sub CopyNumber(ByRef n As Number)
    n = mNumber
End Sub

Friend Property Get IsOne() As Boolean
    If mNumber.Sign <> Positive Then Exit Property
    If mNumber.Precision <> 1 Then Exit Property
    IsOne = mNumber.Digits(0) = 1
End Property

Friend Property Get IsMinusOne() As Boolean
    If mNumber.Sign <> Negative Then Exit Property
    If mNumber.Precision <> 1 Then Exit Property
    IsMinusOne = mNumber.Digits(0) = -1
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFromArray(ByRef Value As Variant)
    If CorArray.IsNull(Value) Then Exit Sub
    
    Dim Bytes() As Byte
    Bytes = Value
    
    Dim l As Long
    l = UBound(Bytes) - LBound(Bytes) + 1
    If l = 0 Then Exit Sub
    
    ' we need to calculate rounding upwards.
    ReDim mNumber.Digits(0 To -Int(-l / 2) - 1)
    
    ' we copy the array incase the source is not 0-based.
    Call CopyMemory(mNumber.Digits(0), Bytes(LBound(Bytes)), l)
    Call Normalize
End Sub

Private Sub InitFromNumber(ByVal Value As Long)
    ReDim mNumber.Digits(0 To 1)
    Call CopyMemory(mNumber.Digits(0), Value, 4)
    Call Normalize
End Sub

Private Sub InitFromFloat(ByVal Value As Double)
    Dim Bytes() As Byte
    ReDim Bytes(0 To 7)
    
    Call CopyMemory(Bytes(0), Value, 8)
    
    Dim isNeg   As Boolean: isNeg = Bytes(7) And &H80
    Dim exp     As Long:    exp = (Bytes(7) And &H7F) * &H10 Or ((Bytes(6) And &HF0) \ &H10)
    If exp = 0 Then Exit Sub
    
    ' remove the sign bit and exponent values, and
    ' add the leading 1 back to the number.
    Bytes(6) = (Bytes(6) And &HF) Or &H10
    Bytes(7) = 0
    
    ReDim mNumber.Digits(0 To 3)
    mNumber.Precision = 4
    Call CopyMemory(mNumber.Digits(0), Bytes(0), 8)
    
    If exp > &H433 Then
        mNumber.Digits = ShiftLeftNumber(mNumber, exp - &H433)
    Else
        mNumber.Digits = ShiftRightNumber(mNumber, &H433 - exp)
    End If

    If isNeg Then Call ApplyTwosComplement(mNumber.Digits)
    Call Normalize
End Sub

Private Sub InitFromCurrency(ByVal Value As Currency)
    Value = Value * 0.0001@
    ReDim mNumber.Digits(0 To 3)
    Call CopyMemory(mNumber.Digits(0), Value, 8)
    Call Normalize
End Sub

Private Sub InitFromDecimal(ByVal Value As Variant)
    Value = Int(Value)
    ReDim mNumber.Digits(0 To 7)
    Call CopyMemory(mNumber.Digits(0), Value, 16)
    
    Dim isNeg As Boolean: isNeg = mNumber.Digits(1) And &H8000
    
    mNumber.Digits(0) = mNumber.Digits(4)
    mNumber.Digits(1) = mNumber.Digits(5)
    mNumber.Digits(4) = mNumber.Digits(2)
    mNumber.Digits(5) = mNumber.Digits(3)
    mNumber.Digits(2) = mNumber.Digits(6)
    mNumber.Digits(3) = mNumber.Digits(7)
    mNumber.Digits(6) = 0
    mNumber.Digits(7) = 0
    
    If isNeg Then Call ApplyTwosComplement(mNumber.Digits)
    Call Normalize
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

Private Function ToDecimalString(ByVal Count As Long) As String
    Const DIGITS_PER_INTEGER As Long = 5
    
    Dim dividend As Number
    dividend = mNumber
    
    Dim Size As Long
    Size = mNumber.Precision * DIGITS_PER_INTEGER
    If Count > Size Then Size = Count
    
    Dim Result As String
    Result = String$(Size + 1, 48)
    
    If mNumber.Sign = Negative Then Call modBigMath.Negate(dividend)
    
    Dim i As Long
    i = Len(Result)
    
    Do While dividend.Precision > 0
        Dim R As Long
        R = SingleInPlaceDivideBy10(dividend)
        If R Then Mid$(Result, i, 1) = Chr$(48 + R)
        i = i - 1
    Loop
    
    If Size - i < Count Then i = Size - Count + 1
    
    If mNumber.Sign = Negative Then
        Mid$(Result, i, 1) = "-"
        i = i - 1
    End If
    
    ToDecimalString = Mid$(Result, i + 1)
End Function

Private Function ToHexString(ByVal Count As Long, Offset As Long) As String
    Const DIGITS_PER_INTEGER As Long = 4
    
    Dim Size As Long
    Size = mNumber.Precision * DIGITS_PER_INTEGER
    If Count > Size Then Size = Count
    
    Dim Result As String
    Result = String$(Size, 48)
    
    Dim c As Long
    Dim k As Long
    Dim i As Long
    For i = 0 To mNumber.Precision - 1
        k = mNumber.Digits(i) And &HFFFF&
        c = Len(Result) - i * DIGITS_PER_INTEGER
        
        Do While k > 0
            If k And &HF Then Mid$(Result, c, 1) = Mid$(HEX_CHARS, (k And &HF) + Offset, 1)
            k = k \ &H10
            c = c - 1
        Loop
    Next i
    
    If Size - c < Count Then
        If mNumber.Sign = Negative Then
            Mid$(Result, Count - Size + 1) = String$(Count - (Size - c), Mid$(HEX_CHARS, &HF + Offset, 1))
        End If
        
        c = Size - Count
    End If
    
    ToHexString = Mid$(Result, c + 1)
End Function

Private Function ToBinaryString(ByVal Count As Long) As String
    Const DIGITS_PER_INTEGER As Long = 16
    
    Dim Size As Long
    Size = mNumber.Precision * DIGITS_PER_INTEGER + 1
    If Count > Size Then Size = Count
    
    Dim Result As String
    Result = String$(Size, 48)
    
    Dim c As Long
    Dim k As Long
    Dim i As Long
    For i = 0 To mNumber.Precision - 1
        k = mNumber.Digits(i) And &HFFFF&
        c = Len(Result) - i * DIGITS_PER_INTEGER
         
        Do While k > 0
            If k And 1 Then Mid$(Result, c, 1) = "1"
            k = k \ &H2
            c = c - 1
        Loop
    Next i
    
    If Size - c < Count Then
        If mNumber.Sign = Negative Then
            Mid$(Result, Count - Size + 1) = String$(Count - (Size - c), 49)
        End If
        
        c = Size - Count
    End If
    
    ToBinaryString = Mid$(Result, c + 1)
End Function

Private Function InternalSetBit(ByVal bit As Long, ByVal Value As Boolean) As BigInteger
    If bit < 0 Then _
        Call Err.Raise(5, , "Bit cannot be negative.")

    Dim Index As Long
    Index = (bit And &HFFFFFFF0) \ &H10
    
    Dim Number() As Integer
    
    If Index < mNumber.Precision Then
        ReDim Number(0 To mNumber.Precision - 1)
        Call CopyMemory(Number(0), mNumber.Digits(0), mNumber.Precision * 2)
    ElseIf (Value = False) And (mNumber.Sign <> Negative) Then
        Set InternalSetBit = Me
        Exit Function
    Else
        ReDim Number(0 To Index)
        Call CopyMemory(Number(0), mNumber.Digits(0), mNumber.Precision * 2)
        
        If mNumber.Sign = Negative Then
            Call FillMemory(Number(mNumber.Precision), (Index - mNumber.Precision + 1) * 2, &HFF)
        End If
    End If
    
    If Value Then
        Number(Index) = Number(Index) Or PowersOf2(bit And &HF)
    Else
        Number(Index) = Number(Index) And (Not PowersOf2(bit And &HF))
    End If
    
    Set InternalSetBit = New BigInteger
    Call InternalSetBit.InternalInit(Number)
End Function

Private Function ParseFormatSpecifier(ByVal Format As String, ByRef Specifier As String) As Long
    If Len(Format) = 0 Then
        Specifier = "d"
        Exit Function
    End If
    
    Select Case Left$(Format, 1)
        Case "x", "X", "b", "B", "d", "D"
            Specifier = Left$(Format, 1)
        Case Else
            Call Err.Raise(5, , "Invalid format specifier")
    End Select
            
    Dim Result  As Long
    Dim i       As Long
    For i = 2 To Len(Format)
        Dim Ch As Long
        Ch = Asc(Mid$(Format, i, 1))
        Select Case Ch
            Case corZeroChar To corNineChar
                Result = Result * 10 + (Ch - 48)
                If Result > 1000000 Then
                    Exit For   ' hopefully over 1 million digits is enough.
                End If
            Case Else
                Call Err.Raise(5, , "Invalid format specifier")
        End Select
    Next i
    
    ParseFormatSpecifier = Result
End Function

Private Function ShiftLeftNumber(ByRef n As Number, ByVal bitCount As Long) As Integer()
    Dim wordCount As Long
    wordCount = (bitCount And &HFFFFFFF0) \ &H10
    
    Dim remainingBits As Long
    remainingBits = bitCount And &HF
    
    Dim newPrecision As Long
    newPrecision = n.Precision + wordCount
    
    Dim newNumber() As Integer
    ReDim newNumber(0 To newPrecision)
    If n.Sign = Negative Then newNumber(newPrecision) = &HFFFF
    
    If remainingBits = 0 Then
        Call CopyMemory(newNumber(wordCount), n.Digits(0), n.Precision * 2)
    Else
        Dim lowermask   As Long: lowermask = ((PowersOf2(16 - remainingBits) And &HFFFF&) - 1) And &HFFFF&
        Dim uppermask   As Long: uppermask = (Not lowermask) And &HFFFF&
        Dim RShift      As Long: RShift = PowersOf2(16 - remainingBits) And &HFFFF&
        Dim LShift      As Long: LShift = PowersOf2(remainingBits) And &HFFFF&
        
        Dim srcInt      As Integer
        Dim dstInt      As Integer
        Dim i           As Long
        Dim j           As Long

        i = wordCount
        Do
            srcInt = n.Digits(j)
            
            If InIDE Then
                newNumber(i) = dstInt Or GetInt(GetLong((srcInt And lowermask)) * LShift)
            Else
                newNumber(i) = dstInt Or ((srcInt And &HFFFF&) And lowermask) * LShift
            End If
            i = i + 1
            
            dstInt = ((srcInt And &HFFFF&) And uppermask) \ RShift
            
            j = j + 1
            If j = n.Precision Then Exit Do
        Loop
        
        If InIDE Then
            If n.Sign = Negative Then dstInt = dstInt Or GetInt(&H10000 - LShift)
        Else
            If n.Sign = Negative Then dstInt = dstInt Or (&H10000 - LShift)
        End If
               
        newNumber(i) = dstInt
    End If

    ShiftLeftNumber = newNumber
End Function

Private Function ShiftRightNumber(ByRef n As Number, ByVal bitCount As Long) As Integer()
    Dim wordCount As Long
    wordCount = (bitCount And &HFFFFFFF0) \ &H10
    
    Dim remainingBits As Long
    remainingBits = bitCount And &HF
    
    Dim newPrecision As Long
    newPrecision = n.Precision - wordCount
    
    Dim newNumber() As Integer
    ReDim newNumber(0 To newPrecision)
    If n.Sign = Negative Then newNumber(newPrecision) = &HFFFF
    
    If remainingBits = 0 Then
        Call CopyMemory(newNumber(0), n.Digits(wordCount), newPrecision * 2)
    Else
        Dim lowermask   As Long: lowermask = ((PowersOf2(remainingBits) And &HFFFF&) - 1) And &HFFFF&
        Dim uppermask   As Long: uppermask = (Not lowermask) And &HFFFF&
        Dim RShift      As Long: RShift = PowersOf2(remainingBits) And &HFFFF&
        Dim LShift      As Long: LShift = PowersOf2(16 - remainingBits) And &HFFFF&
        
        Dim srcInt      As Integer
        Dim dstInt      As Integer
        Dim i           As Long
        Dim j           As Long
        
        j = wordCount
        srcInt = n.Digits(j)
        Do
            dstInt = ((srcInt And &HFFFF&) And uppermask) \ RShift
            
            j = j + 1
            If j = n.Precision Then Exit Do
            
            srcInt = n.Digits(j)
            
            If InIDE Then
                newNumber(i) = dstInt Or GetInt(GetLong(srcInt And lowermask) * LShift)
            Else
                newNumber(i) = dstInt Or (((srcInt And &HFFFF&) And lowermask) * LShift)
            End If
            
            i = i + 1
        Loop
            
        If InIDE Then
            If n.Sign = Negative Then dstInt = dstInt Or GetInt(&H10000 - LShift)
        Else
            If n.Sign = Negative Then dstInt = dstInt Or (&H10000 - LShift)
        End If
        
        newNumber(i) = dstInt
    End If

    ShiftRightNumber = newNumber
End Function


Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Cor.BInt(Value))
End Function
