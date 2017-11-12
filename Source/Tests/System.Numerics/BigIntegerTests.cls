VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigIntegerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BigIntegerTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Numerics"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "NewBigInteger_WhenInitialized_InitializesBigInteger"
            Test.Use 0, NewBytes(), Sign.Zero, 0
            Test.Use 1, NewBytes(1), Sign.Positive, 1
            Test.Use &HFFFFFFFF, NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use &H8000&, NewBytes(0, &H80), Sign.Positive, 2
            Test.Use &H8000, NewBytes(0, &H80), Sign.Negative, 2
            Test.Use &HFF8000, NewBytes(0, &H80, &HFF), Sign.Positive, 3
            Test.Use &HFFFF8000, NewBytes(0, &H80), Sign.Negative, 2
            Test.Use &H7FFF, NewBytes(&HFF, &H7F), Sign.Positive, 2
            Test.Use &HFF7FFF, NewBytes(&HFF, &H7F, &HFF), Sign.Positive, 3
            Test.Use &HFFFF7FFF, NewBytes(&HFF, &H7F, &HFF, &HFF), Sign.Negative, 4
            Test.Use &HFF, NewBytes(&HFF), Sign.Positive, 1
            Test.Use 1#, NewBytes(1), Sign.Positive, 1
            Test.Use -1#, NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use 0#, NewBytes(), Sign.Zero, 0
            Test.Use 0.95, NewBytes(), Sign.Zero, 0
            Test.Use 1@, NewBytes(1), Sign.Positive, 1
            Test.Use -1@, NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use CDec("1"), NewBytes(1), Sign.Positive, 1
            Test.Use CDec("-1"), NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use CDec("12345678901231234567890123.456789"), NewBytes(&HCB, &HE4, &H5F, &H51, &HA2, &H7B, &H22, &H98, &H4C, &H36, &HA), Sign.Positive, 11
            Test.Use NullBytes, NewBytes(), Sign.Zero, 0
            Test.Use NewBytes(), NewBytes(), Sign.Zero, 0
            Test.Use NewBytes(1), NewBytes(1), Sign.Positive, 1
            Test.Use NewBytes(1, 0, 0, 0, 0, 0, 0, 0, 0), NewBytes(1), Sign.Positive, 1
            Test.Use NewBytes(&HFF, &HFF), NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use NewBytes(0, 0, 0, 0, &HFF), NewBytes(0, 0, 0, 0, &HFF), Sign.Positive, 5
            Test.Use NewBytes(0, 0, 0, &H80, &HFF, &HFF), NewBytes(0, 0, 0, &H80), Sign.Negative, 4
            Test.Use NewBytes(0, &H80), NewBytes(0, &H80), Sign.Negative, 2
            
        Case "Parse_WithValue_ReturnsBigInteger"
            Test.Use "12345678901234567890", NewBytes(&HD2, &HA, &H1F, &HEB, &H8C, &HA9, &H54, &HAB), Sign.Positive, 8
            Test.Use "-12345678901234567890", NewBytes(&H2E, &HF5, &HE0, &H14, &H73, &H56, &HAB, &H54, &HFF, &HFF), Sign.Negative, 10
            
        Case "TryParse_WithInvalidValue_ReturnsFalse"
            Test.Use ""
            Test.Use "z"
            Test.Use "1234.123"
            Test.Use "&h"
            Test.Use "0xaa.a"
            Test.Use "0b"
            
        Case "TryParse_WithValidValue_InitializesBigInteger"
            Test.Use "0", NewBytes(), Sign.Zero, 0
            Test.Use "1", NewBytes(&H1), Sign.Positive, 1
            Test.Use "01", NewBytes(&H1), Sign.Positive, 1
            Test.Use "-1", NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use "65535", NewBytes(&HFF, &HFF), Sign.Positive, 2
            Test.Use "  65536  ", NewBytes(0, 0, &H1), Sign.Positive, 3
            Test.Use "  -65536  ", NewBytes(0, 0, &HFF, &HFF), Sign.Negative, 4
            Test.Use "  000000000000000000000255  ", NewBytes(&HFF), Sign.Positive, 1
            Test.Use "12345678901234567890", NewBytes(&HD2, &HA, &H1F, &HEB, &H8C, &HA9, &H54, &HAB), Sign.Positive, 8
            Test.Use "-12345678901234567890", NewBytes(&H2E, &HF5, &HE0, &H14, &H73, &H56, &HAB, &H54, &HFF, &HFF), Sign.Negative, 10
            Test.Use "&h0", NewBytes(), Sign.Zero, 0
            Test.Use "&h00", NewBytes(), Sign.Zero, 0
            Test.Use "&h1", NewBytes(1), Sign.Positive, 1
            Test.Use "0x1", NewBytes(1), Sign.Positive, 1
            Test.Use "&hff", NewBytes(&HFF), Sign.Positive, 1
            Test.Use "0xff", NewBytes(&HFF), Sign.Positive, 1
            Test.Use "0xffff", NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use "0x0ffff", NewBytes(&HFF, &HFF), Sign.Negative, 2
            Test.Use "0x0000ffff", NewBytes(&HFF, &HFF), Sign.Positive, 2
            Test.Use "0x0000ffffffffffffffff", NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF), Sign.Positive, 8
            Test.Use "&h80", NewBytes(&H80), Sign.Positive, 1
            Test.Use "&hff80", NewBytes(&H80, &HFF), Sign.Negative, 2
            Test.Use "&h8000", NewBytes(0, &H80), Sign.Negative, 2
            Test.Use "&h00008000", NewBytes(0, &H80), Sign.Positive, 2
            Test.Use "0x0000ff", NewBytes(&HFF), Sign.Positive, 1
            Test.Use "0xff7fffff", NewBytes(&HFF, &HFF, &H7F, &HFF), Sign.Negative, 4
            Test.Use "0x00000000", NewBytes(), Sign.Zero, 0
            Test.Use "&h80000000", NewBytes(0, 0, 0, &H80), Sign.Negative, 4
            Test.Use "&h800000000000", NewBytes(0, 0, 0, 0, 0, &H80), Sign.Negative, 6
            Test.Use "-&h7fff", NewBytes(&H1, &H80), Sign.Negative, 2
            Test.Use "-&h80000000", NewBytes(0, 0, 0, &H80), Sign.Negative, 4
            Test.Use "0b0", NewBytes(), Sign.Zero, 0
            Test.Use "0b1", NewBytes(1), Sign.Positive, 1
            Test.Use "0b11", NewBytes(3), Sign.Positive, 1
            Test.Use "0b1000", NewBytes(&H8), Sign.Positive, 1
            Test.Use "0b1111", NewBytes(&HF), Sign.Positive, 1
            Test.Use "0b111111111", NewBytes(&HFF, &H1), Sign.Positive, 2
            Test.Use "0b10000000000000000000000000000000", NewBytes(0, 0, 0, &H80), Sign.Negative, 4
            Test.Use "0b000000000000000010000000000000000000000000000000", NewBytes(0, 0, 0, &H80), Sign.Positive, 4
            Test.Use "0b110000000000000000000000000000000", NewBytes(0, 0, 0, &H80, 1), Sign.Positive, 5
            Test.Use "-0b111111100000000", NewBytes(0, &H81), Sign.Negative, 2
        
        Case "ToString_WithNoFormatting_ReturnsDefault"
            Test.Use 0, "0"
            Test.Use 1, "1"
            Test.Use -1, "-1"
            Test.Use 1000000000, "1000000000"
            Test.Use -1000000000, "-1000000000"
            Test.Use &H7FFFFFFF, "2147483647"
            Test.Use &H80000000, "-2147483648"
            
        Case "ToString_WithFormatting_ReturnsFormattedNumber"
            Test.Use 0, "d4", "0000"
            Test.Use &H8000&, "d1", "32768"
            Test.Use &H8000, "d1", "-32768"
            Test.Use &H8000&, "d10", "0000032768"
            Test.Use &H8000&, "d5", "32768"
            Test.Use &H8000&, "d6", "032768"
            Test.Use &H8000, "d0", "-32768"
            Test.Use &H8000, "d6", "-032768"
            Test.Use &H8000, "d5", "-32768"
            Test.Use 0, "x", "0"
            Test.Use 0, "X", "0"
            Test.Use 0, "x4", "0000"
            Test.Use 10, "x", "a"
            Test.Use 10, "X", "A"
            Test.Use 1000000000, "X", "3B9ACA00"
            Test.Use -1000000000, "x", "c4653600"
            Test.Use 2147483647, "x", "7fffffff"
            Test.Use -2147483648#, "x", "80000000"
            Test.Use &HFFFF8000, "x", "8000"
            Test.Use &HFFFF8000, "x5", "f8000"
            Test.Use &HFFFF8000, "x8", "ffff8000"
            Test.Use &HFFFF8000, "x4", "8000"
            Test.Use &HFFFF7FFF, "x", "ffff7fff"
            Test.Use &HFFFF7FFF, "x9", "fffff7fff"
            Test.Use &HFFFF8001, "x", "8001"
            Test.Use &HFFFF8001, "x8", "ffff8001"
            Test.Use &H7FFF, "x4", "7fff"
            Test.Use &H7FFF, "x5", "07fff"
            Test.Use &H7FFF, "x8", "00007fff"
            Test.Use &HFFFF0000, "x", "ffff0000"
            Test.Use &HFFFF0000, "x8", "ffff0000"
            Test.Use 0, "b", "0"
            Test.Use 0, "b4", "0000"
            Test.Use 0, "b8", "00000000"
            Test.Use 1, "b", "1"
            Test.Use 1, "b2", "01"
            Test.Use 1, "b4", "0001"
            Test.Use 3, "b", "11"
            Test.Use 3, "b6", "000011"
            Test.Use 128, "b", "10000000"
            Test.Use &H8080, "b", "1000000010000000"
            Test.Use &H8080, "b20", "11111000000010000000"
            Test.Use -1, "b", "1111111111111111"
            Test.Use -1, "b32", "11111111111111111111111111111111"
            Test.Use &H7FFF, "B", "111111111111111"
            Test.Use &H7FFF, "b32", "00000000000000000111111111111111"
            
        Case "Compare_WithWithValues_ReturnsExpected"
            Test.Use(BInt(0), BInt(0)).Expect Iz.EqualTo(0)
            Test.Use(Nothing, Nothing).Expect Iz.EqualTo(0)
            Test.Use(Nothing, BInt(0)).Expect Iz.EqualTo(0)
            Test.Use(BInt(0), Nothing).Expect Iz.EqualTo(0)
            Test.Use(BInt(1), BInt(1)).Expect Iz.EqualTo(0)
            Test.Use(BInt(1), BInt(0)).Expect Iz.Positive
            Test.Use(BInt(0), BInt(1)).Expect Iz.Negative
            Test.Use(BInt(-1), BInt(-10)).Expect Iz.Positive
            Test.Use(BInt(-10), BInt(-1)).Expect Iz.Negative
            Test.Use(BInt(-128), BInt(-129)).Expect Iz.Positive
            Test.Use(BInt(&H80000000), BInt(&H7FFFFFFF)).Expect Iz.Negative
            Test.Use(BInt(&H7FFFFFFF), BInt(&H80000000)).Expect Iz.Positive
        
        Case "Negate_WithValue_ReturnsExpected"
            Test.Use BInt(0), BInt(0)
            Test.Use BInt(1), BInt(-1)
            Test.Use BInt(-1), BInt(1)
            Test.Use BInt(Nothing), BInt(0)
            Test.Use BInt(255), BInt(-255)
            Test.Use BInt(-255), BInt(255)
            Test.Use BInt(65535), BInt(-65535)
            Test.Use BInt(&H80000000), BInt(2147483648#)
            Test.Use BInt(&H8000), BInt(32768)
        
        Case "Absolute_WithValue_ReturnsExpected"
            Test.Use BInt(0), BInt(0)
            Test.Use BInt(1), BInt(1)
            Test.Use BInt(-1), BInt(1)
            Test.Use BInt(&H80000000), BInt(2147483648#)
            
        Case "Add_WithValues_ReturnsExpected"
            Test.Use 0, 0, 0
            Test.Use 1, 1, 2
            Test.Use &H7FFFFFFF, &H7FFFFFFF, 4294967294#
            Test.Use &H7FFFFFFF, 1, 2147483648#
            Test.Use -1, -1, -2
            Test.Use -1, 1, 0
            Test.Use 1, -1, 0
            Test.Use &H80000000, -1, -2147483649#
            Test.Use -1, &H80000000, -2147483649#
            Test.Use "12345678901234567890", "98765432109876543210", "111111111011111111100"
            
        Case "Subtract_WithValues_ReturnsExpected"
            Test.Use 0, 0, 0
            Test.Use 1, 1, 0
            Test.Use 1, 2, -1
            Test.Use 2, 1, 1
            Test.Use -1, -1, 0
            Test.Use -1, 1, -2
            Test.Use 1, -1, 2
            Test.Use 4294967294#, &H7FFFFFFF, &H7FFFFFFF
            Test.Use 2147483648#, 1, 2147483647
            Test.Use 1, 2147483648#, -2147483647
            Test.Use "111111111011111111100", "12345678901234567890", "98765432109876543210"
            Test.Use "111111111011111111100", "98765432109876543210", "12345678901234567890"
        
        Case "Mulitply_WithValues_ReturnsExpected"
            Test.Use 0, 0, 0
            Test.Use 1, 1, 1
            Test.Use 1, 0, 0
            Test.Use 0, 1, 0
            Test.Use 1, 2, 2
            Test.Use 2, 1, 2
            Test.Use 2147483647, 2, 4294967294#
            Test.Use 2, 2147483647, 4294967294#
            Test.Use "12345678901234567890", "12345678901234567890", "152415787532388367501905199875019052100"
            Test.Use "-12345678901234567890", "-12345678901234567890", "152415787532388367501905199875019052100"
            Test.Use "-12345678901234567890", "12345678901234567890", "-152415787532388367501905199875019052100"
            Test.Use "12345678901234567890", "-12345678901234567890", "-152415787532388367501905199875019052100"
        
        Case "Divide_WithValues_ReturnsExpected"
            Test.Use 0, 1, 0
            Test.Use 4, 2, 2
            Test.Use 2, 4, 0
            Test.Use 0, -1, 0
            Test.Use 4, -2, -2
            Test.Use -4, 2, -2
            Test.Use -4, -2, 2
            Test.Use -4, 4, -1
            Test.Use 4, -4, -1
            Test.Use &H7FFFFFFF, 2, &H3FFFFFFF
            Test.Use "12345678901234567890", "12345678901234567890", 1
            Test.Use "-12345678901234567890", "12345678901234567890", -1
            Test.Use "12345678901234567890", "-12345678901234567890", -1
            
        Case "DevRem_WithValues_ReturnsExpected"
            Test.Use 10, 5, 2, 0
            Test.Use 12, 5, 2, 2
            Test.Use 12, -5, -2, 2
            Test.Use "12345678901234567890", "123456789012345", "100000", "67890"
        
        Case "Pow_WithValues_ReturnsExpected"
            Test.Use 0, 0, 1
            Test.Use 1, 0, 1
            Test.Use 0, 1, 0
            Test.Use 10, 10, "10000000000"
            Test.Use -2, 3, -8
            Test.Use -2, 4, 16
            
        Case "SquareRoot_WithValue_ReturnsExpected"
            Test.Use 4, 2
            Test.Use 15, 3
            Test.Use 100, 10
            Test.Use "123456789012345678901234567", "11111111061111"
            Test.Use "152415787532388367501905199875019052100", "12345678901234567890"
            
        Case "TestBit_WithValue_ReturnsExpected"
            Test.Use 0, 0, False
            Test.Use 0, 15, False
            Test.Use 1, 0, True
            Test.Use 1, 1, False
            Test.Use -1, 0, True
            Test.Use -1, 15, True
            Test.Use -1, 16, True
            Test.Use &HFFFF&, 0, True
            Test.Use &HFFFF&, 15, True
            Test.Use &HFFFF&, 16, False
            Test.Use &HFFFF0000, 15, False
            Test.Use &HFFFF0000, 16, True
            Test.Use &HFFFF0000, 31, True
            Test.Use &HFFFF0000, 32, True
            Test.Use "0x0000ffff0000", 31, True
            Test.Use "0x0000ffff0000", 32, False
            
        Case "SetBit_WithValue_SetsExpectedBit"
            Test.Use 0, 0, &H1
            Test.Use 0, 1, &H2
            Test.Use 0, 2, &H4
            Test.Use 1, 7, &H81
            Test.Use 0, 16, &H10000
            Test.Use 0, 31, &H80000000
            Test.Use 0, 32, "0x100000000"
            
        Case "ClearBit_WithValue_ClearsExpectedBit"
            Test.Use &HFFFF, 0, &HFFFE
            Test.Use &HFFFF, 15, &H7FFF
            Test.Use &HFFFFFFFF, 16, &HFFFEFFFF
            Test.Use &HFFFFFFFF, 31, &H7FFFFFFF
            Test.Use &H7FFFFFFF, 32, &H7FFFFFFF
            
        Case "AndBits_WithValue_ReturnsAndedValue"
            Test.Use 0, 0, 0
            Test.Use &HFF00, &HFF, 0
            Test.Use &HFF, &HFF00, 0
            Test.Use &HFFFF, &HFF0, &HFF0
            Test.Use &HFFFF0000, &H1FFFF, &H10000
            Test.Use &H1FFFF, &HFFFF0000, &H10000
            Test.Use &HFFFF0000, &HFFFF, &HFFFF0000
            Test.Use &HFFFF, &HFFFF0000, &HFFFF0000
            
        Case "AndNotBits_WithValue_ReturnsAndNottedValue"
            Test.Use 0, 0, 0
            Test.Use 0, &HFFFF, 0
            Test.Use &HFFFF, &HFFFF, 0
            Test.Use &HFFFF, 0, &HFFFF
        
        Case "OrBits_WithValue_ReturnsOredValue"
            Test.Use 0, 0, 0
            Test.Use &HFF00, &HFF, &HFFFF
            Test.Use &HFF, &HFF00, &HFFFF
            Test.Use &H80000000, &H1, &H80000001
            Test.Use &H1, &H80000000, &H80000001
        
        Case "XorBits_WithValue_ReturnsXoredValue"
            Test.Use 0, 0, 0
            Test.Use 0, &HFFFF, &HFFFF
            Test.Use &HFFFF, 0, &HFFFF
            Test.Use &HFFFF, &HFFFF, 0
            Test.Use &HFF00, &HFF, &HFFFF
            Test.Use &HFF00, &HFFFF, &HFF
            Test.Use &HFFFF, &HFF00, &HFF
            Test.Use &HFFFF0000, &H7FFF, &HFFFF7FFF
            Test.Use &H7FFF, &HFFFF0000, &HFFFF7FFF
            Test.Use &HFFFF, &HFFFF8000, &H7FFF
            
        Case "NotBits_WhenCalled_ReturnsNottedValue"
            Test.Use 0, -1
            Test.Use -1, 0
            Test.Use 1, &HFFFE
            Test.Use &HFFFF0000, &HFFFF&
            
        Case "Precision_WithValue_ReturnsExpected"
            Test.Use 0, 0
            Test.Use 1, 1
            Test.Use &HFF, 1
            Test.Use &H1FF, 2
            Test.Use &HFFFFFF, 3
            Test.Use &HFFFF, 2
            Test.Use &HFFFFFFFF, 2
            
    End Select
End Sub

Public Sub BigInteger_WhenCreated_InitializesToZero()
    Dim Number As New BigInteger
    
    AssertNumber Number, NewBytes(), Sign.Zero, 0
End Sub

Public Sub NewBigInteger_WhenInitialized_InitializesBigInteger(ByRef Value As Variant, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Sign, ByVal ExpectedPrecision As Long)
    Dim Number As BigInteger
    
    Set Number = NewBigInteger(Value)
    
    AssertNumber Number, ExpectedBytes, ExpectedSign, ExpectedPrecision
End Sub

Public Sub Parse_WithValue_ReturnsBigInteger(ByVal s As String, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Sign, ByVal ExpectedPrecision As Long)
    Dim Number As BigInteger
    
    Set Number = BigInteger.Parse(s)
    
    AssertNumber Number, ExpectedBytes, ExpectedSign, ExpectedPrecision
End Sub

Public Sub Parse_WithInvalidValue_ThrowsFormatException()
    On Error Resume Next
    
    BigInteger.Parse "hello"
    
    AssertFormatException Err
End Sub

Public Sub TryParse_WithInvalidValue_ReturnsFalse(ByVal s As String)
    Dim Result As BigInteger
    Dim Actual As Boolean
    
    Actual = BigInteger.TryParse(s, Result)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub TryParse_WithValidValue_InitializesBigInteger(ByVal s As String, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Sign, ByVal ExpectedPrecision As Long)
    Dim Number As BigInteger
    Dim Actual As Boolean
    
    Actual = BigInteger.TryParse(s, Number)
    
    Assert.That Actual, Iz.True
    AssertNumber Number, ExpectedBytes, ExpectedSign, ExpectedPrecision
End Sub

Public Sub ToString_WithNoFormatting_ReturnsDefault(ByVal Value As Long, ByVal Expected As String)
    Dim Number As BigInteger
    Dim Actual As String
    
    Set Number = NewBigInteger(Value)
    
    Actual = Number.ToString
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFormatting_ReturnsFormattedNumber(ByVal Value As Long, ByVal Format As String, ByVal Expected As String)
    Dim Number As BigInteger
    Dim Actual As String
    
    Set Number = NewBigInteger(Value)
    
    Actual = Number.ToString(Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Function Compare_WithWithValues_ReturnsExpected(ByVal x As BigInteger, ByVal y As BigInteger) As Long
    Dim Actual As Long
    
    Actual = BigInteger.Compare(x, y)
    
    Compare_WithWithValues_ReturnsExpected = Actual
End Function

Public Sub Negate_WithValue_ReturnsExpected(ByVal Number As BigInteger, ByVal Expected As BigInteger)
    Dim Actual As BigInteger
    
    Set Actual = Number.Negate
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Negate_WhenNegativeToPositive_ShouldSetPrecision()
    Dim Number As BigInteger
    Set Number = BInt("-12345678901234567890")
    
    Dim Actual As Long
    Actual = Number.Negate.Precision
    
    Assert.That Actual, Iz.EqualTo(8)
End Sub

Public Sub Absolute_WithValue_ReturnsExpected(ByVal Number As BigInteger, ByVal Expected As BigInteger)
    Dim Actual As BigInteger
    
    Set Actual = Number.Absolute
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Add_WithValues_ReturnsExpected(ByRef LValue As Variant, ByRef RValue As Variant, ByRef Expected As Variant)
    Dim LHS As BigInteger
    Dim RHS As BigInteger
    
    Set LHS = BInt(LValue)
    Set RHS = BInt(RValue)
    
    Dim Actual As BigInteger
    Set Actual = LHS.Add(RHS)
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub Subtract_WithValues_ReturnsExpected(ByRef LValue As Variant, ByRef RValue As Variant, ByRef Expected As Variant)
    Dim LHS As BigInteger
    Dim RHS As BigInteger
    
    Set LHS = BInt(LValue)
    Set RHS = BInt(RValue)
    
    Dim Actual As BigInteger
    Set Actual = LHS.Subtract(RHS)
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub Mulitply_WithValues_ReturnsExpected(ByRef LValue As Variant, ByRef RValue As Variant, ByRef Expected As Variant)
    Dim LHS As BigInteger
    Dim RHS As BigInteger
    
    Set LHS = BInt(LValue)
    Set RHS = BInt(RValue)
    
    Dim Actual As BigInteger
    Set Actual = LHS.Multiply(RHS)
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub Divide_WithValues_ReturnsExpected(ByRef DividendValue As Variant, ByRef DivisorValue As Variant, ByRef Expected As Variant)
    Dim Dividend As BigInteger
    Dim Divisor As BigInteger
    
    Set Dividend = BInt(DividendValue)
    Set Divisor = BInt(DivisorValue)
    
    Dim Actual As BigInteger
    Set Actual = Dividend.Divide(Divisor)
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub DevRem_WithValues_ReturnsExpected(ByRef DividendValue As Variant, ByRef DivisorValue As Variant, ByRef ExpectedQuotient As Variant, ByRef ExpectedRemainder As Variant)
    Dim Dividend As BigInteger
    Dim Divisor As BigInteger
    Dim Remainder As BigInteger
    
    Set Dividend = BInt(DividendValue)
    Set Divisor = BInt(DivisorValue)
    
    Dim Actual As BigInteger
    Set Actual = Dividend.DivRem(Divisor, Remainder)
    
    Assert.That Actual, Equals(BInt(ExpectedQuotient))
    Assert.That Remainder, Equals(BInt(ExpectedRemainder))
End Sub


Public Sub Factorial_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    BInt(-1).Factorial
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Factorial_WithValue_ReturnsExpected()
    Dim Actual As BigInteger
    
    Set Actual = BInt(4).Factorial
    
    Assert.That Actual, Equals(BInt(24))
End Sub

Public Sub Pow_WithValues_ReturnsExpected(ByVal Value As Long, ByVal Exponent As Long, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value).Pow(BInt(Exponent))
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub Pow_WithNegativeExponent_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    BigInteger.Pow BInt(2), BInt(-2)
    
    AssertArgumentOutOfRangeException Err, "Exponent"
End Sub

Public Sub SquareRoot_WithValue_ReturnsExpected(ByRef Value As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value).SquareRoot
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub ShiftRight_WhenShiftingZeroBits_ReturnsOriginalValue()
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Ten.ShiftRight(0)
    
    Assert.That Actual, Equals(BigInteger.Ten)
End Sub

Public Sub ShiftRight_WhenShiftingForPositiveValue_ShiftsBits()
    Dim Actual As BigInteger
    
    Set Actual = BInt(&H7FFFFFFF).ShiftRight(1)
    
    Assert.That Actual, Equals(BInt(&H3FFFFFFF))
End Sub

Public Sub ShiftRight_WithNegativeCount_ShiftsLeft()
    Dim Actual As BigInteger
    
    Set Actual = BInt(1).ShiftRight(-1)
    
    Assert.That Actual, Equals(BInt(2))
End Sub

Public Sub ShiftRight_WhenShiftingAllBitsForPositiveValue_ReturnsZero()
    Dim Actual As BigInteger
    
    Set Actual = BInt(&H7FFFFFFF).ShiftRight(32)
    
    Assert.That Actual, Equals(BigInteger.Zero)
End Sub

Public Sub ShiftRight_WhenShiftingAllBitsForNegativeValue_ReturnsMinusOne()
    Dim Actual As BigInteger
    
    Set Actual = BInt(&H80000000).ShiftRight(128)
    
    Assert.That Actual, Equals(BigInteger.MinusOne)
End Sub

Public Sub ShiftLeft_WhenShiftingZeroBits_ReturnsOriginalValue()
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Ten.ShiftLeft(0)
    
    Assert.That Actual, Equals(BigInteger.Ten)
End Sub

Public Sub ShiftLeft_WhenShiftingForPositiveValue_ShiftsBits()
    Dim Actual As BigInteger
    
    Set Actual = BInt(&H3FFFFFFF).ShiftLeft(1)
    
    Assert.That Actual, Equals(BInt(&H7FFFFFFE))
End Sub

Public Sub ShiftLeft_WithNegativeCount_ShiftsRight()
    Dim Actual As BigInteger
    
    Set Actual = BInt(&H7FFFFFFF).ShiftLeft(-1)
    
    Assert.That Actual, Equals(BInt(&H3FFFFFFF))
End Sub

Public Sub ShiftLeft_WhenShiftingAllBits_ShiftsBits()
    Dim Actual As BigInteger
    
    Set Actual = BInt(&HFFFF&).ShiftLeft(16)
    
    Assert.That Actual, Equals(BInt("0x0000ffff0000"))
End Sub

Public Sub ShiftLeft_WhenShiftingNegativeOne_ReturnsMinus65536()
    Dim Actual As BigInteger
    
    Set Actual = BInt(-1).ShiftLeft(16)
    
    Assert.That Actual, Equals(BInt(&HFFFF0000))
End Sub

Public Sub TestBit_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Bit As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = BInt(Value).TestBit(Bit)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SetBit_WithValue_SetsExpectedBit(ByRef Value As Variant, ByVal Bit As Long, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value).SetBit(Bit)
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub ClearBit_WithValue_ClearsExpectedBit(ByRef Value As Variant, ByVal Bit As Long, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value).ClearBit(Bit)
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub AndBits_WithValue_ReturnsAndedValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value1).AndBits(BInt(Value2))
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub AndNotBits_WithValue_ReturnsAndNottedValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value1).AndNotBits(BInt(Value2))
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub OrBits_WithValue_ReturnsOredValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value1).OrBits(BInt(Value2))
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub XorBits_WithValue_ReturnsXoredValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value1).XorBits(BInt(Value2))
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub NotBits_WhenCalled_ReturnsNottedValue(ByRef Value As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value).NotBits
    
    Assert.That Actual, Equals(BInt(Expected))
End Sub

Public Sub Rnd_WhenCalled_ReturnsRandomNumber()
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Rnd(20)
    
    Assert.That Actual, NotEquals(BigInteger.Zero)
End Sub

Public Sub Rnd_WithSameSeed_ProducesSameNumber()
    Dim Number1 As BigInteger
    Dim Number2 As BigInteger
    
    BigInteger.SeedRnd -13
    Set Number1 = BigInteger.Rnd(20)
    BigInteger.SeedRnd -13
    Set Number2 = BigInteger.Rnd(20)
    
    Assert.That Number1, Equals(Number2)
End Sub

Public Sub Rnd_WithNegative_ReturnsNegativeRandomNumber()
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Rnd(20, True)
    
    Assert.That Actual.Sign, Iz.EqualTo(Sign.Negative)
End Sub

Public Sub Precision_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = BInt(Value).Precision
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Private Sub AssertNumber(ByVal Number As BigInteger, ByRef ExpectedBytes() As Byte, ByRef ExpectedSign As Sign, ByRef ExpectedPrecision As Long)
    Assert.That Number.ToByteArray, Iz.EqualTo(ExpectedBytes)
    Assert.That Number.Sign, Iz.EqualTo(ExpectedSign)
    Assert.That Number.Precision, Iz.EqualTo(ExpectedPrecision)
End Sub
