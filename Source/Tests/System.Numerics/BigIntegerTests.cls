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
        Case "NewBigInteger_WithLong_InitializesBigInteger"
            Test.Use 0, NewBytes(0), 0
            Test.Use 1, NewBytes(1), 1
            Test.Use 8, NewBytes(8), 1
            Test.Use -8, NewBytes(&HF8), -1
            Test.Use &H80, NewBytes(&H80, 0), 1
            Test.Use &H8000, NewBytes(0, &H80), -1
            Test.Use &H80000000, NewBytes(0, 0, 0, &H80), -1
            Test.Use &HFF, NewBytes(&HFF, 0), 1
            Test.Use &HFF01, NewBytes(1, &HFF), -1
            Test.Use &HFFF, NewBytes(&HFF, &HF), 1
            Test.Use -1, NewBytes(&HFF), -1
            Test.Use &HFFFF&, NewBytes(&HFF, &HFF, 0), 1
            Test.Use &H1FFFF, NewBytes(&HFF, &HFF, 1), 1
            Test.Use &HFFFFF, NewBytes(&HFF, &HFF, &HF), 1
            Test.Use &HFFFFFFFF, NewBytes(&HFF), -1
            
        Case "NewBigInteger_WithByteArray_InitializesBigInteger"
            Test.Use NewBytes(), NewBytes(0), 0
            Test.Use NewBytes(&H80), NewBytes(&H80), -1
            Test.Use NewBytes(&HFF), NewBytes(&HFF), -1
            Test.Use NewBytes(&HFF, &HFF), NewBytes(&HFF), -1
            Test.Use NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H4, &HF1, &H67, &HA8, &HA5, &H67, &H68), NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H4, &HF1, &H67, &HA8, &HA5, &H67, &H68), 1
            
        Case "NewBigInteger_WithDouble_InitializesBigInteger"
            Test.Use 0#, NewBytes(0), 0
            Test.Use 0.95, NewBytes(0), 0
            Test.Use 1#, NewBytes(1), 1
            Test.Use 1.9, NewBytes(1), 1
            Test.Use 1.1, NewBytes(1), 1
            Test.Use -1#, NewBytes(&HFF), -1
            Test.Use 2#, NewBytes(2), 1
            Test.Use -2#, NewBytes(&HFE), -1
            Test.Use 1E+55, NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H4, &HF1, &H67, &HA8, &HA5, &H67, &H68), 1
            Test.Use -1E+55, NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &HFC, &HE, &H98, &H57, &H5A, &H98, &H97), -1
            Test.Use 1E+16, NewBytes(&H0, &H0, &HC1, &H6F, &HF2, &H86, &H23), 1
            Test.Use -1E+16, NewBytes(&H0, &H0, &H3F, &H90, &HD, &H79, &HDC), -1
            Test.Use 1E-16, NewBytes(0), 0
            Test.Use -1E-16, NewBytes(0), 0
            Test.Use -4294901760#, NewBytes(&H0, &H0, &H1, &H0, &HFF), -1
            
        Case "NewBigInteger_WithCurrency_InitializesBigInteger"
            Test.Use 0@, NewBytes(0), 0
            Test.Use 1@, NewBytes(1), 1
            Test.Use -1@, NewBytes(&HFF), -1
            Test.Use 1.1@, NewBytes(1), 1
            Test.Use -1.1@, NewBytes(&HFF), -1
            Test.Use 0.9999@, NewBytes(0), 0
            Test.Use -0.9999@, NewBytes(0), 0
        
        Case "NewBigInteger_WithDecimal_InitializesBigInteger"
            Test.Use 1, NewBytes(1), 1
            Test.Use -1, NewBytes(&HFF), -1
            Test.Use "12345678901231234567890123.456789", NewBytes(&HCB, &HE4, &H5F, &H51, &HA2, &H7B, &H22, &H98, &H4C, &H36, &HA), 1
            Test.Use "-12345678901231234567890123.456789", NewBytes(&H35, &H1B, &HA0, &HAE, &H5D, &H84, &HDD, &H67, &HB3, &HC9, &HF5), -1
            
        Case "Parse_WithValidValueForNumberStylesInteger_ReturnsBigInteger"
            Test.Use "0", NewBytes(0), 0
            Test.Use "1", NewBytes(&H1), 1
            Test.Use "01", NewBytes(&H1), 1
            Test.Use "-1", NewBytes(&HFF), -1
            Test.Use " -1", NewBytes(&HFF), -1
            Test.Use "-1 ", NewBytes(&HFF), -1
            Test.Use " 1", NewBytes(&H1), 1
            Test.Use "1 ", NewBytes(&H1), 1
            Test.Use "+1", NewBytes(&H1), 1
            Test.Use " +1", NewBytes(&H1), 1
            Test.Use "+1 ", NewBytes(&H1), 1
            Test.Use "255", NewBytes(&HFF, 0), 1
            Test.Use "65535", NewBytes(&HFF, &HFF, 0), 1
            Test.Use "  65536  ", NewBytes(0, 0, &H1), 1
            Test.Use "  -65536  ", NewBytes(0, 0, &HFF), -1
            Test.Use "  000000000000000000000255  ", NewBytes(&HFF, 0), 1
            Test.Use "  -000000000000000000000255  ", NewBytes(&H1, &HFF), -1
            Test.Use "12345678901234567890", NewBytes(&HD2, &HA, &H1F, &HEB, &H8C, &HA9, &H54, &HAB, 0), 1
            Test.Use "-12345678901234567890", NewBytes(&H2E, &HF5, &HE0, &H14, &H73, &H56, &HAB, &H54, &HFF), -1
            
        Case "Parse_WithInvalidValueForNumberStylesInteger_ThrowsFormatException"
            Test.Use "1 2"
            Test.Use "- 12"
            Test.Use "12-"
            Test.Use "F"
            Test.Use "(1)"
            Test.Use "1,000"
            Test.Use "1.0"
            Test.Use "+ 12"
            Test.Use "12+"
            Test.Use "$12"
            Test.Use "1e+10"
            
        Case "Parse_WithInvalidValueForNumberStylesAllowParentheses_ThrowsFormatException"
            Test.Use "-1"
            Test.Use " (1)"
            Test.Use "(1) "
            Test.Use "( 1)"
            Test.Use "(1 )"
            Test.Use "(1"
            Test.Use "1)"
            Test.Use "-(1)"
            Test.Use "(-1)"
            
        
            
        Case "ToString_WithNoFormatting_ReturnsDefault"
            Test.Use 0, "0"
            Test.Use 1, "1"
            Test.Use -1, "-1"
            Test.Use 1000000000, "1000000000"
            Test.Use -1000000000, "-1000000000"
            Test.Use &H7FFFFFFF, "2147483647"
            Test.Use &H80000000, "-2147483648"
            Test.Use &HFFFF0000, "-65536"
            
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
            
'        Case "Compare_WithWithValues_ReturnsExpected"
'            Test.Use(BInt(0), BInt(0)).Expect Iz.EqualTo(0)
'            Test.Use(Nothing, Nothing).Expect Iz.EqualTo(0)
'            Test.Use(Nothing, BInt(0)).Expect Iz.EqualTo(0)
'            Test.Use(BInt(0), Nothing).Expect Iz.EqualTo(0)
'            Test.Use(BInt(1), BInt(1)).Expect Iz.EqualTo(0)
'            Test.Use(BInt(1), BInt(0)).Expect Iz.Positive
'            Test.Use(BInt(0), BInt(1)).Expect Iz.Negative
'            Test.Use(BInt(-1), BInt(-10)).Expect Iz.Positive
'            Test.Use(BInt(-10), BInt(-1)).Expect Iz.Negative
'            Test.Use(BInt(-128), BInt(-129)).Expect Iz.Positive
'            Test.Use(BInt(&H80000000), BInt(&H7FFFFFFF)).Expect Iz.Negative
'            Test.Use(BInt(&H7FFFFFFF), BInt(&H80000000)).Expect Iz.Positive
'
'        Case "Negate_WithValue_ReturnsExpected"
'            Test.Use BInt(0), BInt(0)
'            Test.Use BInt(1), BInt(-1)
'            Test.Use BInt(-1), BInt(1)
'            Test.Use BInt(Nothing), BInt(0)
'            Test.Use BInt(255), BInt(-255)
'            Test.Use BInt(-255), BInt(255)
'            Test.Use BInt(65535), BInt(-65535)
'            Test.Use BInt(&H80000000), BInt(2147483648#)
'            Test.Use BInt(&H8000), BInt(32768)

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
            Test.Use "4294901760", 31, True
            Test.Use "4294901760", 32, False
            
        Case "SetBit_WithValue_SetsExpectedBit"
            Test.Use 0, 0, NewBytes(1), 1
            Test.Use 0, 15, NewBytes(0, &H80, 0), 1
            Test.Use 0, 16, NewBytes(0, 0, 1), 1
            Test.Use 1, 0, NewBytes(1), 1
            Test.Use 1, 1, NewBytes(3), 1
            Test.Use &H10000, 0, NewBytes(1, 0, 1), 1
            Test.Use 0, 31, NewBytes(0, 0, 0, &H80, 0), 1
            Test.Use 0, 32, NewBytes(0, 0, 0, 0, 1), 1
            Test.Use 1, 16, NewBytes(1, 0, 1), 1
            Test.Use -1, 0, NewBytes(&HFF), -1
            Test.Use -1, 16, NewBytes(&HFF), -1
            
        Case "ClearBit_WithValue_ClearsExpectedBit"
            Test.Use &HFFFF, 0, NewBytes(&HFE), -1
            Test.Use &HFFFF, 15, NewBytes(&HFF, &H7F, &HFF), -1
            Test.Use &HFFFFFFFF, 16, NewBytes(&HFF, &HFF, &HFE), -1
            Test.Use &HFFFFFFFF, 31, NewBytes(&HFF, &HFF, &HFF, &H7F, &HFF), -1
            Test.Use &HFFFFFFFF, 32, NewBytes(&HFF, &HFF, &HFF, &HFF, &HFE), -1
            Test.Use &H7FFFFFFF, 32, NewBytes(&HFF, &HFF, &HFF, &H7F), 1
            Test.Use 0, 0, NewBytes(0), 0
            Test.Use 1, 0, NewBytes(0), 0
            Test.Use 2, 0, NewBytes(&H2), 1
            Test.Use 3, 0, NewBytes(&H2), 1
            Test.Use &H10001, 0, NewBytes(&H0, &H0, &H1), 1
            Test.Use &H8001&, 0, NewBytes(&H0, &H80, 0), 1
            Test.Use &H8001&, 15, NewBytes(1), 1
            Test.Use &H8000&, 15, NewBytes(0), 0
            Test.Use &H10001, 16, NewBytes(1), 1
            Test.Use 1, 16, NewBytes(1), 1
            Test.Use -1, 0, NewBytes(&HFE), -1
            Test.Use -1, 7, NewBytes(&H7F, &HFF), -1
            Test.Use -1, 15, NewBytes(&HFF, &H7F, &HFF), -1
            Test.Use -1, 16, NewBytes(&HFF, &HFF, &HFE), -1
            
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
            
        Case "NotBits_WhenCalled_ReturnsExpected"
            Test.Use 0, NewBytes(&HFF), -1
            Test.Use -1, NewBytes(0), 0
            Test.Use 1, NewBytes(&HFE), -1
            Test.Use &HFFFF0000, NewBytes(&HFF, &HFF, &H0), 1
            Test.Use &HFFFF&, NewBytes(&H0, &H0, &HFF), -1
                       
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use 0, 0, True
            Test.Use 1, 1, True
            Test.Use 1, 2, False
            Test.Use 2, 1, False
            Test.Use -1, -1, True
            Test.Use 1, -1, False
            Test.Use -1, 1, False
            Test.Use &H7FFFFFFF, &H7FFFFFFF, True
            Test.Use &H7FFFFFFF, &H80000000, False
            Test.Use &HFFFFFFFF, &H80000000, False
            Test.Use &H80000000, &HFFFFFFFF, False
            Test.Use &H80000000, &H80000000, True
            Test.Use &H80000001, 1, False
            Test.Use 1, &H80000001, False
            Test.Use -123456789012345#, -123456789012345#, True
            Test.Use -123456789012345#, 123456789012345#, False
            Test.Use 123456789012345#, -123456789012345#, False
            Test.Use 0, Nothing, True
            Test.Use Nothing, 0, True
            Test.Use Nothing, Nothing, True
            
    End Select
End Sub

Public Sub Setup()
    ClearException
End Sub

Public Sub BigInteger_WhenCreated_InitializesToZero()
    Dim Number As New BigInteger
    
    AssertNumber Number, NewBytes(0), 0
End Sub

Public Sub NewBigInteger_WhenInitialized_InitializesBigInteger(ByRef Value As Variant, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Sign, ByVal ExpectedPrecision As Long)
    Dim Number As BigInteger
    
    Set Number = NewBigInteger(Value)
    
    AssertNumber Number, ExpectedBytes, ExpectedSign
End Sub

Public Sub NewBigInteger_WithNullArray_ThrowsArgumentNullException()
    On Error Resume Next
    
    NewBigInteger NullBytes
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub Parse_WithValidValueForNumberStylesInteger_ReturnsBigInteger(ByVal s As String, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Parse(s, NumberStyles.Integer)
    
    AssertNumber Actual, ExpectedBytes, ExpectedSign
End Sub

Public Sub Parse_WithInvalidValueForNumberStylesInteger_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    BigInteger.Parse s, NumberStyles.Integer
    
    AssertFormatException Err
End Sub

Public Sub Parse_WithValidValueForNumberStylesAllowParentheses_ReturnsBigInteger()
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Parse("(1)", NumberStyles.AllowParentheses)
    
    AssertNumber Actual, NewBytes(&HFF), -1
End Sub

Public Sub Parse_WithInvalidValueForNumberStylesAllowParentheses_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    BigInteger.Parse s, NumberStyles.AllowParentheses
    
    AssertFormatException Err
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

Public Sub ShiftRight_WithOne_ReturnsZero()
    Dim Actual As BigInteger
    
    Set Actual = BInt(1).ShiftRight(1)
    
    Assert.That Actual.IsZero, Iz.True
End Sub

Public Sub ShiftRight_WhenAlreadyShiftedAllBits_ReturnsZero()
    Dim Actual As BigInteger
    
    Set Actual = BInt(1).ShiftRight(1).ShiftRight(1)
    
    Assert.That Actual.IsZero, Iz.True
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
    
    Assert.That Actual, Equals(BInt(NewBytes(0, 0, &HFF, &HFF, 0)))
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

Public Sub SetBit_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    BInt(1).SetBit -1
    
    AssertArgumentOutOfRangeException Err, "Bit"
End Sub

Public Sub SetBit_WithValue_SetsExpectedBit(ByRef Value As Variant, ByVal Bit As Long, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value).SetBit(Bit)
    
    Assert.That Actual.ToByteArray, Iz.EqualTo(ExpectedBytes), "Wrong byte array"
    Assert.That Actual.Sign, Iz.EqualTo(ExpectedSign), "Wrong sign"
End Sub

Public Sub ClearBit_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    BInt(1).ClearBit -1
    
    AssertArgumentOutOfRangeException Err, "Bit"
End Sub

Public Sub ClearBit_WithValue_ClearsExpectedBit(ByRef Value As Variant, ByVal Bit As Long, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = BInt(Value).ClearBit(Bit)
    
    Assert.That Actual.ToByteArray, Iz.EqualTo(ExpectedBytes), "Wrong byte array"
    Assert.That Actual.Sign, Iz.EqualTo(ExpectedSign), "Wrong sign"
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

Public Sub NotBits_WhenCalled_ReturnsExpected(ByRef Value As Variant, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual() As Byte
    Dim ActualSign As Long
    Dim Number As BigInteger
    
    Set Number = NewBigInteger(Value).NotBits
    Actual = Number.ToByteArray
    ActualSign = Number.Sign
    
    Assert.That Actual, Iz.EqualTo(ExpectedBytes)
    Assert.That ActualSign, Iz.EqualTo(ExpectedSign)
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

Public Sub NewBigInteger_WithLong_InitializesBigInteger(ByVal Value As Long, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual() As Byte
    Dim ActualSign As Long
    
    Dim Number As BigInteger
    Set Number = NewBigInteger(Value)
    
    Actual = Number.ToByteArray
    ActualSign = Number.Sign
    
    Assert.That Actual, Iz.EqualTo(ExpectedBytes)
    Assert.That ActualSign, Iz.EqualTo(ExpectedSign)
End Sub

Public Sub NewBigInteger_WithDouble_InitializesBigInteger(ByVal Value As Double, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual() As Byte
    Dim ActualSign As Long
    
    Dim Number As BigInteger
    Set Number = NewBigInteger(Value)
    
    Actual = Number.ToByteArray
    ActualSign = Number.Sign
    
    Assert.That Actual, Iz.EqualTo(ExpectedBytes)
    Assert.That ActualSign, Iz.EqualTo(ExpectedSign)
End Sub

Public Sub NewBigInteger_WithByteArray_InitializesBigInteger(ByRef Value() As Byte, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual() As Byte
    Dim ActualSign As Long
    
    Dim Number As BigInteger
    Set Number = NewBigInteger(Value)
    
    Actual = Number.ToByteArray
    ActualSign = Number.Sign
    
    Assert.That Actual, Iz.EqualTo(ExpectedBytes)
    Assert.That ActualSign, Iz.EqualTo(ExpectedSign)
End Sub

Public Sub NewBigInteger_WithCurrency_InitializesBigInteger(ByVal Value As Currency, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual() As Byte
    Dim ActualSign As Long
    
    Dim Number As BigInteger
    Set Number = NewBigInteger(Value)
    
    Actual = Number.ToByteArray
    ActualSign = Number.Sign
    
    Assert.That Actual, Iz.EqualTo(ExpectedBytes)
    Assert.That ActualSign, Iz.EqualTo(ExpectedSign)
End Sub

Public Sub NewBigInteger_WithDecimal_InitializesBigInteger(ByRef Value As Variant, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual() As Byte
    Dim ActualSign As Long
    
    Dim Number As BigInteger
    Set Number = NewBigInteger(CDec(Value))
    
    Actual = Number.ToByteArray
    ActualSign = Number.Sign
    
    Assert.That Actual, Iz.EqualTo(ExpectedBytes)
    Assert.That ActualSign, Iz.EqualTo(ExpectedSign)
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByRef a As Variant, ByRef b As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = BigInteger.Equals(BInt(a), BInt(b))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithNothing_ReturnsTrue()
    Dim Actual As Boolean
    
    Actual = BigInteger.Zero.Equals(Nothing)
    
    Assert.That Actual, Iz.True
End Sub
