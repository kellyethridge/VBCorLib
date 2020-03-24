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
            
        Case "NewBigInteger_WithInt64_InitializesBigInteger"
            Test.Use Int64.Zero, NewBytes(0), 0
            Test.Use Int64.One, NewBytes(1), 1
            Test.Use Int64.MinusOne, NewBytes(&HFF), -1
            Test.Use CInt64(&H8000), NewBytes(0, &H80), -1
            Test.Use CInt64(-&H800000), NewBytes(0, 0, &H80), -1
            Test.Use CInt64(&H80000000), NewBytes(0, 0, 0, &H80), -1
            Test.Use CInt64("0xFFFF"), NewBytes(&HFF, &HFF, 0), 1
            Test.Use CInt64("0xFFFFFF"), NewBytes(&HFF, &HFF, &HFF, 0), 1
            Test.Use CInt64("0xFFFFFFFF"), NewBytes(&HFF, &HFF, &HFF, &HFF, 0), 1
            Test.Use CInt64("0xFFFFFFFFFF"), NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, 0), 1
            Test.Use CInt64("0xffffffffffffffff"), NewBytes(&HFF), -1
            Test.Use CInt64("0x8000000000000000"), NewBytes(0, 0, 0, 0, 0, 0, 0, &H80), -1
            Test.Use CInt64("0xffff800000000000"), NewBytes(0, 0, 0, 0, 0, &H80), -1
            Test.Use CInt64("0x7fffffffffffffff"), NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &H7F), 1
            
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
            ' Unspecified Format
            Test.Use 0, "", "0"
            Test.Use 1, "", "1"
            Test.Use -1, "", "-1"
            Test.Use 100, "", "100"
            Test.Use &H7FFFFFFF, "", "2147483647"
            Test.Use &H80000000, "", "-2147483648"
            
            ' General Format (.NET ignores g formatting rules and just displays the whole number)
            Test.Use 0, "g", "0"
            Test.Use 1, "g", "1"
            Test.Use -1, "g", "-1"
            Test.Use &H7FFFFFFF, "g", "2147483647"
            Test.Use &H80000000, "g", "-2147483648"
            Test.Use 0, "g0", "0"
            Test.Use 1, "g0", "1"
            Test.Use -1, "g0", "-1"
            Test.Use &H7FFFFFFF, "g0", "2147483647"
            Test.Use &H80000000, "g0", "-2147483648"
            Test.Use 0, "g5", "0"
            Test.Use 1, "g5", "1"
            Test.Use -1, "g5", "-1"
            Test.Use &H7FFFFFFF, "g5", "2147483647"
            Test.Use &H80000000, "g5", "-2147483648"
            Test.Use 12345, "ggg", "ggg"
        
            ' Currency Format
            Test.Use 1, "c", "$1.00"
            Test.Use 1, "c0", "$1"
            Test.Use 1, "c1", "$1.0"
            Test.Use 1, "c3", "$1.000"
            Test.Use -1, "c", "($1.00)"
            Test.Use 12345, "c", "$12,345.00"
            Test.Use "12345678901234567890", "c4", "$12,345,678,901,234,567,890.0000"
            
            ' Decimal Format
            Test.Use 1, "d", "1"
            Test.Use 1, "d1", "1"
            Test.Use 1, "d2", "01"
            Test.Use -1, "d2", "-01"
            
            ' Exponent Format
            Test.Use 1, "e", "1.000000e+000"
            Test.Use 1, "e0", "1e+000"
            Test.Use 1, "e1", "1.0e+000"
            Test.Use 10, "e", "1.000000e+001"
            Test.Use 10, "e0", "1e+001"
            Test.Use 10, "e1", "1.0e+001"
            Test.Use -1, "e", "-1.000000e+000"
            Test.Use 19, "e0", "2e+001"
            Test.Use 12345, "e", "1.234500e+004"
            Test.Use -12345, "e4", "-1.2345e+004"
            Test.Use 12345, "e3", "1.235e+004"
            Test.Use 12345, "E1", "1.2E+004"
            Test.Use 0, "e", "0.000000e+000"
            Test.Use 123456789, "e4", "1.2346e+008"
            
            ' Fixed Format
            Test.Use 0, "f", "0.00"
            Test.Use 1, "f", "1.00"
            Test.Use 100, "f", "100.00"
            Test.Use 1234, "f", "1234.00"
            Test.Use -1234, "f", "-1234.00"
            Test.Use 1234, "f4", "1234.0000"
            Test.Use 1234, "f0", "1234"
            
            ' Number Format
            Test.Use 0, "n", "0.00"
            Test.Use 1, "n", "1.00"
            Test.Use 100, "n", "100.00"
            Test.Use 1234, "n", "1,234.00"
            Test.Use 1234, "n0", "1,234"
            Test.Use -1234, "n", "-1,234.00"
            Test.Use 1234, "n4", "1,234.0000"
            Test.Use 123, "n", "123.00"
            
            ' Percentage Format
            Test.Use 0, "p", "0.00 %"
            Test.Use 1, "p", "100.00 %"
            Test.Use 156, "p", "15,600.00 %"
            
            ' Hex Format
            Test.Use 0, "X", "0"
            Test.Use 0, "x2", "00"
            Test.Use 1, "x", "1"
            Test.Use 1, "x2", "01"
            Test.Use 255, "x", "0ff"
            Test.Use 255, "X", "0FF"
            Test.Use 255, "X4", "00FF"
            Test.Use -1, "x", "f"
            Test.Use -1, "X", "F"
            Test.Use -1, "x5", "fffff"
            Test.Use &H8000, "x", "8000"
            Test.Use &H1111, "x", "1111"
            Test.Use &HA, "x", "0a"
            Test.Use &HFFF80000, "x", "80000"
            Test.Use &HFFF80000, "x8", "fff80000"
            
        Case "Compare_WithWithValues_ReturnsExpected"
            Test.Use(CBigInt(0), CBigInt(0)).Expect Iz.EqualTo(0)
            Test.Use(Nothing, Nothing).Expect Iz.EqualTo(0)
            Test.Use(Nothing, CBigInt(0)).Expect Iz.EqualTo(0)
            Test.Use(CBigInt(0), Nothing).Expect Iz.EqualTo(0)
            Test.Use(CBigInt(1), CBigInt(1)).Expect Iz.EqualTo(0)
            Test.Use(CBigInt(1), CBigInt(0)).Expect Iz.Positive
            Test.Use(CBigInt(0), CBigInt(1)).Expect Iz.Negative
            Test.Use(CBigInt(-1), CBigInt(-10)).Expect Iz.Positive
            Test.Use(CBigInt(-10), CBigInt(-1)).Expect Iz.Negative
            Test.Use(CBigInt(-128), CBigInt(-129)).Expect Iz.Positive
            Test.Use(CBigInt(&H80000000), CBigInt(&H7FFFFFFF)).Expect Iz.Negative
            Test.Use(CBigInt(&H7FFFFFFF), CBigInt(&H80000000)).Expect Iz.Positive

        Case "Negate_WithValue_ReturnsExpected"
            Test.Use 0, 0
            Test.Use 1, -1
            Test.Use -1, 1
            Test.Use Nothing, 0
            Test.Use 255, -255
            Test.Use -255, 255
            Test.Use 65535, -65535
            Test.Use -65535, 65535
            Test.Use &H80000000, 2147483648#
            Test.Use &H8000, 32768

        Case "Absolute_WithValue_ReturnsExpected"
            Test.Use CBigInt(0), CBigInt(0)
            Test.Use CBigInt(1), CBigInt(1)
            Test.Use CBigInt(-1), CBigInt(1)
            Test.Use CBigInt(&H80000000), CBigInt(2147483648#)
            
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
            Test.Use "-12345678901234567890", "2", "-24691357802469135780"
            Test.Use "2", "-12345678901234567890", "-24691357802469135780"
                    
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
            Test.Use "123456789012345678901234567", "11111111061111", "11111111061111"
            
        Case "DivRem_WithValues_ReturnsExpected"
            Test.Use 10, 5, 2, 0
            Test.Use 12, 5, 2, 2
            Test.Use 12, -5, -2, -2
            Test.Use "12345678901234567890", "123456789012345", "100000", "67890"
        
        Case "Modulus_WithValues_ReturnsExpected"
            Test.Use 10, 5, 0
            Test.Use 10, 3, 1
            Test.Use 0, 2, 0
            Test.Use -5, 2, -1
            Test.Use 5, -2, 1
            Test.Use -5, -2, -1
        
        Case "Pow_WithValues_ReturnsExpected"
            Test.Use 0, 0, 1
            Test.Use 1, 0, 1
            Test.Use 0, 1, 0
            Test.Use 10, 10, "10000000000"
            Test.Use -2, 3, -8
            Test.Use -2, 4, 16
            Test.Use 2, 100, "1267650600228229401496703205376"
                        
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
            
        Case "BitwiseAnd_WithValue_ReturnsAndedValue"
            Test.Use 0, 0, 0
            Test.Use &HFF00, &HFF, 0
            Test.Use &HFF, &HFF00, 0
            Test.Use &HFFFF, &HFF0, &HFF0
            Test.Use &HFFFF0000, &H1FFFF, &H10000
            Test.Use &H1FFFF, &HFFFF0000, &H10000
            Test.Use &HFFFF0000, &HFFFF, &HFFFF0000
            Test.Use &HFFFF, &HFFFF0000, &HFFFF0000
            Test.Use "72057594037927935", &HFF, &HFF
            
        Case "BitwiseAndNot_WithValue_ReturnsAndNottedValue"
            Test.Use 0, 0, 0
            Test.Use 0, &HFFFF, 0
            Test.Use &HFFFF, &HFFFF, 0
            Test.Use &HFFFF, 0, &HFFFF
        
        Case "BitwiseOr_WithValue_ReturnsOredValue"
            Test.Use 0, 0, 0
            Test.Use &HFF00, &HFF, &HFFFF
            Test.Use &HFF, &HFF00, &HFFFF
            Test.Use &H80000000, &H1, &H80000001
            Test.Use &H1, &H80000000, &H80000001
            Test.Use &HFFFF0000, &HFFFF, -1
            Test.Use &HFFFF, &HFFFF0000, -1
            Test.Use &HFF0000, &HFFFF, -1
            Test.Use &HFFFF, &HFF0000, -1
            Test.Use &HFFFF&, &HFFFF, -1
            Test.Use &HFFFF, &HFFFF&, -1
            
        Case "BitwiseXor_WithValue_ReturnsXoredValue"
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
            
        Case "BitwiseNot_WhenCalled_ReturnsExpected"
            Test.Use 0, -1
            Test.Use -1, 0
            Test.Use 1, -2
            Test.Use &HFFFF0000, 65535
            Test.Use &HFFFF&, -65536
            Test.Use &HFF00, &HFF
            Test.Use &HFF, &HFF00
            Test.Use -255, 254
            Test.Use 254, -255
            Test.Use &H7FFFFFFF, &H80000000
                       
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
            
        Case "IsEven_WithValue_ReturnsExpected"
            Test.Use 0, True
            Test.Use 1, False
            Test.Use -1, False
            Test.Use 2, True
            Test.Use -2, True
            
        Case "GreatestCommonDivisor_WithValues_ReturnsExpected"
            Test.Use 10, 5, 5
            Test.Use 5, 10, 5
            Test.Use -10, 5, 5
            Test.Use -10, -5, 5
            Test.Use 10, 3, 1
            Test.Use 0, 5, 5
            Test.Use 5, 0, 5
            Test.Use 0, 0, 0
            
    End Select
End Sub

Public Sub Setup()
    ClearException
End Sub

Public Sub BigInteger_WhenCreated_InitializesToZero()
    Dim Number As New BigInteger
    
    AssertNumber Number, NewBytes(0), 0
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

Public Sub ToString_WithFormatting_ReturnsFormattedNumber(ByVal Value As Variant, ByVal Format As String, ByVal Expected As String)
    Dim Number As BigInteger
    Dim Actual As String
    Set Number = CBigInt(Value)
        
    Actual = Number.ToString(Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Function Compare_WithWithValues_ReturnsExpected(ByVal x As BigInteger, ByVal y As BigInteger) As Long
    Dim Actual As Long
    
    Actual = BigInteger.Compare(x, y)
    
    Compare_WithWithValues_ReturnsExpected = Actual
End Function

Public Sub Negate_WithValue_ReturnsExpected(ByRef Value As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value).Negate
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub Negate_WhenNegativeToPositive_ShouldSetPrecision()
    Dim Number As BigInteger
    Set Number = CBigInt("-12345678901234567890")
    
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
    
    Set LHS = CBigInt(LValue)
    Set RHS = CBigInt(RValue)
    
    Dim Actual As BigInteger
    Set Actual = LHS.Add(RHS)
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub Subtract_WithValues_ReturnsExpected(ByRef LValue As Variant, ByRef RValue As Variant, ByRef Expected As Variant)
    Dim LHS As BigInteger
    Dim RHS As BigInteger
    
    Set LHS = CBigInt(LValue)
    Set RHS = CBigInt(RValue)
    
    Dim Actual As BigInteger
    Set Actual = LHS.Subtract(RHS)
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub Mulitply_WithValues_ReturnsExpected(ByRef LValue As Variant, ByRef RValue As Variant, ByRef Expected As Variant)
    Dim LHS As BigInteger
    Dim RHS As BigInteger
    
    Set LHS = CBigInt(LValue)
    Set RHS = CBigInt(RValue)
    
    Dim Actual As BigInteger
    Set Actual = LHS.Multiply(RHS)
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub Divide_WithValues_ReturnsExpected(ByRef DividendValue As Variant, ByRef DivisorValue As Variant, ByRef Expected As Variant)
    Dim Dividend As BigInteger
    Dim Divisor As BigInteger
    
    Set Dividend = CBigInt(DividendValue)
    Set Divisor = CBigInt(DivisorValue)
    
    Dim Actual As BigInteger
    Set Actual = Dividend.Divide(Divisor)
    
    Assert.That Actual, Equals(CBigInt(Expected))
    Assert.That Dividend, Equals(CBigInt(DividendValue))
    Assert.That Divisor, Equals(CBigInt(DivisorValue))
End Sub

Public Sub DivRem_WithValues_ReturnsExpected(ByRef DividendValue As Variant, ByRef DivisorValue As Variant, ByRef ExpectedQuotient As Variant, ByRef ExpectedRemainder As Variant)
    Dim Dividend As BigInteger
    Dim Divisor As BigInteger
    Dim Remainder As BigInteger
    
    Set Dividend = CBigInt(DividendValue)
    Set Divisor = CBigInt(DivisorValue)
    
    Dim Actual As BigInteger
    Set Actual = Dividend.DivRem(Divisor, Remainder)
    
    Assert.That Actual, Equals(CBigInt(ExpectedQuotient))
    Assert.That Remainder, Equals(CBigInt(ExpectedRemainder))
End Sub


Public Sub Factorial_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CBigInt(-1).Factorial
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Factorial_WithValue_ReturnsExpected()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(4).Factorial
    
    Assert.That Actual, Equals(CBigInt(24))
End Sub

Public Sub Pow_WithValues_ReturnsExpected(ByVal Value As Long, ByVal Exponent As Long, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value).Pow(Exponent)
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub Pow_WithNegativeExponent_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    BigInteger.Pow CBigInt(2), -2
    
    AssertArgumentOutOfRangeException Err, "Exponent"
End Sub

Public Sub SquareRoot_WithValue_ReturnsExpected(ByRef Value As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger

    Set Actual = CBigInt(Value).SquareRoot

    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub ShiftRight_WhenShiftingZeroBits_ReturnsOriginalValue()
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Ten.ShiftRight(0)
    
    Assert.That Actual, Equals(BigInteger.Ten)
End Sub

Public Sub ShiftRight_WhenShiftingForPositiveValue_ShiftsBits()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(&H7FFFFFFF).ShiftRight(1)
    
    Assert.That Actual, Equals(CBigInt(&H3FFFFFFF))
End Sub

Public Sub ShiftRight_WithNegativeCount_ShiftsLeft()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(1).ShiftRight(-1)
    
    Assert.That Actual, Equals(CBigInt(2))
End Sub

Public Sub ShiftRight_WhenShiftingAllBitsForPositiveValue_ReturnsZero()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(&H7FFFFFFF).ShiftRight(32)
    
    Assert.That Actual, Equals(BigInteger.Zero)
End Sub

Public Sub ShiftRight_WithOne_ReturnsZero()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(1).ShiftRight(1)
    
    Assert.That Actual.IsZero, Iz.True
End Sub

Public Sub ShiftRight_WhenAlreadyShiftedAllBits_ReturnsZero()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(1).ShiftRight(1).ShiftRight(1)
    
    Assert.That Actual.IsZero, Iz.True
End Sub

Public Sub ShiftRight_WhenShiftingAllBitsForNegativeValue_ReturnsMinusOne()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(&H80000000).ShiftRight(128)
    
    Assert.That Actual, Equals(BigInteger.MinusOne)
End Sub

Public Sub ShiftLeft_WhenShiftingZeroBits_ReturnsOriginalValue()
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Ten.ShiftLeft(0)
    
    Assert.That Actual, Equals(BigInteger.Ten)
End Sub

Public Sub ShiftLeft_WhenShiftingForPositiveValue_ShiftsBits()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(&H3FFFFFFF).ShiftLeft(1)
    
    Assert.That Actual, Equals(CBigInt(&H7FFFFFFE))
End Sub

Public Sub ShiftLeft_WithNegativeCount_ShiftsRight()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(&H7FFFFFFF).ShiftLeft(-1)
    
    Assert.That Actual, Equals(CBigInt(&H3FFFFFFF))
End Sub

Public Sub ShiftLeft_WhenShiftingAllBits_ShiftsBits()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(&HFFFF&).ShiftLeft(16)
    
    Assert.That Actual, Equals(CBigInt(NewBytes(0, 0, &HFF, &HFF, 0)))
End Sub

Public Sub ShiftLeft_WhenShiftingNegativeOne_ReturnsMinus65536()
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(-1).ShiftLeft(16)
    
    Assert.That Actual, Equals(CBigInt(&HFFFF0000))
End Sub

Public Sub TestBit_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Bit As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = CBigInt(Value).TestBit(Bit)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SetBit_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CBigInt(1).SetBit -1
    
    AssertArgumentOutOfRangeException Err, "Bit"
End Sub

Public Sub SetBit_WithValue_SetsExpectedBit(ByRef Value As Variant, ByVal Bit As Long, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value).SetBit(Bit)
    
    Assert.That Actual.ToByteArray, Iz.EqualTo(ExpectedBytes), "Wrong byte array"
    Assert.That Actual.Sign, Iz.EqualTo(ExpectedSign), "Wrong sign"
End Sub

Public Sub ClearBit_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CBigInt(1).ClearBit -1
    
    AssertArgumentOutOfRangeException Err, "Bit"
End Sub

Public Sub ClearBit_WithValue_ClearsExpectedBit(ByRef Value As Variant, ByVal Bit As Long, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value).ClearBit(Bit)
    
    Assert.That Actual.ToByteArray, Iz.EqualTo(ExpectedBytes), "Wrong byte array"
    Assert.That Actual.Sign, Iz.EqualTo(ExpectedSign), "Wrong sign"
End Sub

Public Sub BitwiseAnd_WithValue_ReturnsAndedValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value1).BitwiseAnd(CBigInt(Value2))
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub BitwiseAndNot_WithValue_ReturnsAndNottedValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value1).BitwiseAndNot(CBigInt(Value2))
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub BitwiseOr_WithValue_ReturnsOredValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value1).BitwiseOr(CBigInt(Value2))
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub BitwiseXor_WithValue_ReturnsXoredValue(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value1).BitwiseXor(CBigInt(Value2))
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub BitwiseNot_WhenCalled_ReturnsExpected(ByRef Value As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = CBigInt(Value).BitwiseNot
    
    Assert.That Actual, Equals(CBigInt(Expected))
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
    
    Actual = CBigInt(Value).Precision
    
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
    
    Actual = BigInteger.Equals(CBigInt(a), CBigInt(b))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithNothing_ReturnsTrue()
    Dim Actual As Boolean
    
    Actual = BigInteger.Zero.Equals(Nothing)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Divide_WithDivisorOfZero_ThrowsDivideByZeroException()
    On Error Resume Next
    
    CBigInt(10).Divide BigInteger.Zero
    
    AssertDivideByZeroException Err
End Sub

Public Sub DivRem_WithDivisorOfZero_ThrowsDivideByZeroException()
    On Error Resume Next
    
    CBigInt(10).DivRem BigInteger.Zero, Nothing
    
    AssertDivideByZeroException Err
End Sub

Public Sub Modulus_WithValues_ReturnsExpected(ByRef DividendValue As Variant, ByRef DivisorValue As Variant, ByRef ExpectedValue As Variant)
    Dim Dividend As BigInteger
    Dim Divisor As BigInteger
    Dim Actual As BigInteger
    
    Set Dividend = CBigInt(DividendValue)
    Set Divisor = CBigInt(DivisorValue)
    
    Set Actual = Dividend.Modulus(Divisor)
    
    Assert.That Actual, Equals(CBigInt(ExpectedValue))
End Sub

Public Sub IsEven_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = CBigInt(Value).IsEven
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GreatestCommonDivisor_WithValues_ReturnsExpected(ByRef Left As Variant, ByRef Right As Variant, ByRef Expected As Variant)
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.GreatestCommonDivisor(CBigInt(Left), CBigInt(Right))
    
    Assert.That Actual, Equals(CBigInt(Expected))
End Sub

Public Sub NewBigInteger_WithInt64_InitializesBigInteger(ByRef Value As Variant, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = Cor.NewBigInteger(Value)
    
    Assert.That Actual.ToByteArray, Iz.EqualTo(ExpectedBytes), "#1"
    Assert.That Actual.Sign, Iz.EqualTo(ExpectedSign)
End Sub
