VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Int64Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Int64Tests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use NewInt64(1, 1), NewInt64(1, 1), True
            Test.Use NewInt64(0, 1), NewInt64(1, 1), False
            Test.Use NewInt64(1, 0), NewInt64(1, 1), False
        
        Case "Compare_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, Int64.Zero, 0
            Test.Use Int64.Zero, Int64.MinusOne, 1
            Test.Use Int64.Zero, Int64.One, -1
            Test.Use Int64.One, Int64.Zero, 1
            Test.Use Int64.One, Int64.MinusOne, 1
            Test.Use Int64.One, Int64.One, 0
            Test.Use Int64.MinusOne, Int64.One, -1
            Test.Use Int64.MinusOne, Int64.MinusOne, 0
            Test.Use Int64.MinusOne, Int64.Zero, -1
            Test.Use NewInt64(&H80000000, 0), Int64.Zero, 1
            Test.Use NewInt64(&H80000000, 0), Int64.MinusOne, 1
            Test.Use NewInt64(0, &H80000000), Int64.MinusOne, -1
            Test.Use NewInt64(&H80000000, 0), Int64.One, 1
            Test.Use NewInt64(&H80000000, 0), NewInt64(&H80000000, 0), 0
            Test.Use Int64.Zero, NewInt64(&H80000000, 0), -1
            Test.Use NewInt64(&H80000000, 0), NewInt64(&H80000001, 0), -1
            Test.Use NewInt64(&H80000001, 0), NewInt64(&H80000000, 0), 1
            
        Case "Sign_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, 0
            Test.Use Int64.MinusOne, -1
            Test.Use Int64.One, 1
            
        Case "Negate_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, Int64.Zero
            Test.Use Int64.One, Int64.MinusOne
            Test.Use NewInt64(10, 10), NewInt64(&HFFFFFFF6, &HFFFFFFF5)
            
        Case "Absolute_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, Int64.Zero
            Test.Use Int64.One, Int64.One
            Test.Use Int64.MinusOne, Int64.One
            Test.Use NewInt64(&HFFFFFFFF, 0), NewInt64(&HFFFFFFFF, 0)
            Test.Use NewInt64(0, &HFFFFFFFF), NewInt64(0, 1)
            
        Case "Add_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, Int64.Zero, Int64.Zero
            Test.Use Int64.Zero, Int64.One, Int64.One
            Test.Use NewInt64(&HFFFFFFFF, 0), Int64.One, NewInt64(0, &H1)
            Test.Use Int64.One, NewInt64(&HFFFFFFFF, 0), NewInt64(0, &H1)
            Test.Use NewInt64(&H7FFFFFFF, 0), Int64.One, NewInt64(&H80000000, &H0)
            Test.Use Int64.One, NewInt64(&H7FFFFFFF, 0), NewInt64(&H80000000, &H0)
            Test.Use NewInt64(&HFFFFFFFF, 0), NewInt64(&HFFFFFFFF, 0), NewInt64(&HFFFFFFFE, &H1)
            Test.Use Int64.MinusOne, Int64.One, Int64.Zero
            Test.Use Int64.One, Int64.MinusOne, Int64.Zero
            Test.Use Int64.MinusOne, Int64.MinusOne, NewInt64(&HFFFFFFFE, &HFFFFFFFF)
            Test.Use NewInt64(0, 1), NewInt64(&H80000000, &HFFFFFFFF), NewInt64(&H80000000, 0)
            Test.Use NewInt64(&H80000000, &HFFFFFFFF), NewInt64(0, 1), NewInt64(&H80000000, 0)
            
        Case "Subract_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, Int64.Zero, Int64.Zero
            Test.Use Int64.One, Int64.One, Int64.Zero
            Test.Use Int64.MinusOne, Int64.MinusOne, Int64.Zero
            Test.Use Int64.MinusOne, Int64.One, NewInt64(&HFFFFFFFE, &HFFFFFFFF)
            Test.Use Int64.One, Int64.MinusOne, NewInt64(2, 0)
            Test.Use NewInt64(&H0, &H1), Int64.One, NewInt64(&HFFFFFFFF, &H0)
            Test.Use NewInt64(&HFFFFFFFF, &H1), Int64.One, NewInt64(&HFFFFFFFE, &H1)
            Test.Use NewInt64(&H0, &H1), NewInt64(&HFFFFFFFF, 0), Int64.One
            Test.Use NewInt64(&H0, &HFFFFFFFF), NewInt64(&H1, 0), NewInt64(&HFFFFFFFF, &HFFFFFFFE)
            
        Case "Multiply_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, Int64.Zero, Int64.Zero
            Test.Use Int64.One, Int64.Zero, Int64.Zero
            Test.Use Int64.Zero, Int64.One, Int64.Zero
            Test.Use Int64.One, Int64.MinusOne, Int64.MinusOne
            Test.Use Int64.MinusOne, Int64.One, Int64.MinusOne
            Test.Use Int64.MinusOne, Int64.MinusOne, Int64.One
            Test.Use NewInt64(&H7FFFFFFF, 0), NewInt64(&H7FFFFFFF, 0), NewInt64(&H1, &H3FFFFFFF)
            Test.Use NewInt64(&HFFFFFFFF, &H3FFFFFFF), NewInt64(2, 0), NewInt64(&HFFFFFFFE, &H7FFFFFFF)
            Test.Use NewInt64(2, 0), NewInt64(&HFFFFFFFF, &H3FFFFFFF), NewInt64(&HFFFFFFFE, &H7FFFFFFF)
            
        Case "Divide_WithValues_ReturnsExpected"
            Test.Use Int64.One, Int64.One, Int64.One
            Test.Use Int64.MinusOne, Int64.One, Int64.MinusOne
            Test.Use NewInt64(1234, 0), NewInt64(2, 0), NewInt64(617, 0)
            Test.Use NewInt64(&H90123456, &H12345678), NewInt64(5, 0), NewInt64(&H5003A411, &H3A4114B)
            Test.Use NewInt64(&H90123456, &H12345678), NewInt64(0, 5), NewInt64(&H3A4114B, &H0)
            Test.Use Int64.MinusOne, Int64.MinusOne, Int64.One
            Test.Use NewInt64(&H90123456, &H12345678), NewInt64(0, &HFFFFFFFB), NewInt64(&HFC5BEEB5, &HFFFFFFFF)
            Test.Use Int64.Zero, Int64.One, Int64.Zero
        
        Case "DivRem_WithValues_ReturnsExpected"
            Test.Use Int64.One, Int64.One, Int64.Zero
            Test.Use Int64.MinusOne, Int64.One, Int64.Zero
            Test.Use NewInt64(1234, 0), NewInt64(2, 0), Int64.Zero
            Test.Use NewInt64(&H90123456, &H12345678), NewInt64(5, 0), NewInt64(1, 0)
            Test.Use NewInt64(&H90123456, &H12345678), NewInt64(0, 5), NewInt64(&H90123456, &H1)
            Test.Use Int64.MinusOne, Int64.MinusOne, Int64.Zero
            Test.Use NewInt64(&H90123456, &H12345678), NewInt64(0, &HFFFFFFFB), NewInt64(&H90123456, &H1)
            Test.Use Int64.Zero, Int64.One, Int64.Zero
            
        Case "ShiftLeft_WithValues_ReturnsExpected"
            Test.Use Int64.One, 0, Int64.One
            Test.Use Int64.One, 1, NewInt64(2, 0)
            Test.Use NewInt64(&H80000001, 0), 1, NewInt64(2, 1)
            Test.Use NewInt64(&H80000001, 0), 2, NewInt64(4, 2)
            Test.Use Int64.One, 32, NewInt64(0, 1)
            Test.Use Int64.One, 63, NewInt64(0, &H80000000)
            Test.Use Int64.One, 64, Int64.Zero
            Test.Use NewInt64(0, &H40000000), 1, NewInt64(0, &H80000000)
            Test.Use NewInt64(&HFFFFFFFF, &H0), 16, NewInt64(&HFFFF0000, &HFFFF&)
            
        Case "ShiftRight_WithValues_ReturnsExpected"
            Test.Use Int64.One, 0, Int64.One
            Test.Use Int64.One, 1, Int64.Zero
            Test.Use NewInt64(&H80000000, 0), 1, NewInt64(&H40000000, 0)
            Test.Use NewInt64(0, 1), 1, NewInt64(&H80000000, 0)
            Test.Use NewInt64(0, &H80000000), 64, Int64.Zero
            Test.Use NewInt64(0, &H80000000), 1, NewInt64(0, &H40000000)
            Test.Use NewInt64(0, &H80000000), 63, NewInt64(1, 0)
            Test.Use NewInt64(&H0, &HFFFFFFFF), 16, NewInt64(&HFFFF0000, &HFFFF&)
        
        Case "ToString_WithValues_ReturnsExpected"
            Test.Use Int64.Zero, "0"
            Test.Use Int64.One, "1"
            Test.Use Int64.MinusOne, "-1"
            Test.Use NewInt64(10, 0), "10"
            Test.Use NewInt64(100, 0), "100"
            Test.Use NewInt64(1000, 0), "1000"
            Test.Use NewInt64(10000, 0), "10000"
            Test.Use NewInt64(100000, 0), "100000"
            Test.Use NewInt64(1000000, 0), "1000000"
            Test.Use NewInt64(10000000, 0), "10000000"
            Test.Use NewInt64(100000000, 0), "100000000"
            Test.Use NewInt64(1000000000, 0), "1000000000"
            Test.Use Int64.MaxValue, "9223372036854775807"
            Test.Use Int64.MinValue, "-9223372036854775808"
            Test.Use NewInt64(&HFFFFFFFF, 0), "4294967295"
            
        Case "ToString_WithFormatting_ReturnsExpected"
            Test.Use Int64.One, "f2", "1.00"
            Test.Use Int64.One, "c", "$1.00"
            Test.Use Int64.MinusOne, "c", "($1.00)"
            Test.Use NewInt64(&H90123456, &H12345678), "X", "1234567890123456"
            Test.Use NewInt64(&HF, &HF), "X16", "0000000F0000000F"
            Test.Use NewInt64(49, 0), "$", "1"
            
        Case "TryParse_WithValidValues_ReturnsExpected"
            Test.Use "0", Int64.Zero
            Test.Use "1", Int64.One
            Test.Use "-1", Int64.MinusOne
            Test.Use "-9223372036854775808", Int64.MinValue
            Test.Use "9223372036854775807", Int64.MaxValue
        
        Case "TryParseEx_WithAllowTrailingWhiteSpace_ReturnsExpected"
            Test.Use "0", Int64.Zero
            Test.Use "1", Int64.One
            Test.Use "01", Int64.One
            Test.Use "255", NewInt64(255, 0)
            Test.Use "65535", NewInt64(&HFFFF&, 0)
            Test.Use "1 ", Int64.One
            Test.Use "255 ", NewInt64(255, 0)
            Test.Use "65535   ", NewInt64(&HFFFF&, 0)
            
        Case "ParseEx_WithInvalidValueAndAllowTrailingWhiteSpace_ThrowsFormatException"
            Test.Use " 1"
            Test.Use "-1"
            Test.Use "- 1"
            Test.Use " -1"
            Test.Use "(1)"
            Test.Use "+ 1"
            Test.Use "+1"
            Test.Use " +1"
            Test.Use "$1"
            Test.Use "1$"
            Test.Use "1e+10"
            Test.Use "FF"
            Test.Use "1.0"
            Test.Use "1,000"
            Test.Use "1 1"
            
        Case "TryParseEx_WithAllowCurrencySymbol_ReturnsExpected"
            Test.Use "0", Int64.Zero
            Test.Use "1", Int64.One
            Test.Use "$1", Int64.One
            Test.Use "1$", Int64.One
            
        Case "ParseEx_WithInvalidValueAndAllowCurrencySymbol_ThrowsFormatException"
            Test.Use "-1"
            Test.Use " 1"
            Test.Use "1 "
            Test.Use "(1)"
            Test.Use "+1"
            Test.Use "1e+10"
            Test.Use "FF"
            Test.Use "1.0"
            Test.Use "1,000"
            Test.Use "1 1"
            Test.Use "1$1"
            Test.Use "$1$"
            Test.Use "$$1"
            Test.Use "1$$"
            
        Case "TryParseEx_WithAllowDecimalPoint_ReturnsExpected"
            Test.Use "0", Int64.Zero
            Test.Use "1", Int64.One
            Test.Use "1.0", Int64.One
            Test.Use "0.0", Int64.Zero
            Test.Use "1.", Int64.One
            Test.Use ".0", Int64.Zero
        
        Case "ParseEx_WithInvalidValueAndAllowDecimalPoint_ThrowsFormatException"
            Test.Use "-1"
            Test.Use " 1"
            Test.Use "1 "
            Test.Use "(1)"
            Test.Use "+1"
            Test.Use "$1"
            Test.Use "1$"
            Test.Use "1 -"
            Test.Use "1e+10"
            Test.Use "FF"
            Test.Use "1,000"
            Test.Use "1 1"
            Test.Use "1+1"
            Test.Use "1-1"
            Test.Use "1.1"
            Test.Use "1..0"
            Test.Use "++1"
            Test.Use "1++"
            Test.Use "--1"
            Test.Use "1--"
        
        Case "TryParseEx_WithAllowExponent_ReturnsExpected"
            Test.Use "1E0", Int64.One
            Test.Use "1e2", NewInt64(&H64, 0)
            Test.Use "1e10", NewInt64(&H540BE400, &H2)
            Test.Use "1e+10", NewInt64(&H540BE400, &H2)
            
        Case "ParseEx_WithInvalidValueAndAllowExponent_ThrowsFormatException"
            Test.Use "-1"
            Test.Use " 1"
            Test.Use "1 "
            Test.Use "+1"
            Test.Use "$1"
            Test.Use "1$"
            Test.Use "1 -"
            Test.Use "1e-10"
            Test.Use "FF"
            Test.Use "1,000"
            Test.Use "1 1"
            Test.Use "1+1"
            Test.Use "1-1"
            Test.Use "1.1"
            Test.Use "1..0"
            Test.Use "++1"
            Test.Use "1++"
            Test.Use "--1"
            Test.Use "1--"
            Test.Use ".0"
            Test.Use " (1) "
            Test.Use "( 1)"
            Test.Use "(1 )"
            Test.Use "1)"
            Test.Use "1e"
            Test.Use "1e+"
            Test.Use "1e-"
            Test.Use "1e10e"
        
        Case "TryParseEx_WithAllowHexSpecifier_ReturnsExpected"
            Test.Use "0", Int64.Zero
            Test.Use "00000000", Int64.Zero
            Test.Use "1", Int64.One
            Test.Use "7f", NewInt64(&H7F, 0)
            Test.Use "100", NewInt64(&H100, 0)
            Test.Use "f", NewInt64(&HF, 0)
            Test.Use "ffffffff", NewInt64(&HFFFFFFFF, 0)
            Test.Use "ffffffffffffffff", Int64.MinusOne
            Test.Use "8000000000000000", Int64.MinValue
            Test.Use "7fffffffffffffff", Int64.MaxValue
            
        Case "TryParse_WithHexPrefix_ReturnsExpected"
            Test.Use "&h0", Int64.Zero
            Test.Use "&h00000000", Int64.Zero
            Test.Use "&h1", Int64.One
            Test.Use "0x7f", NewInt64(&H7F, 0)
            Test.Use "0x100", NewInt64(&H100, 0)
            Test.Use "0Xf", NewInt64(&HF, 0)
            Test.Use "0Xffffffff", NewInt64(&HFFFFFFFF, 0)
            Test.Use "0xffffffffffffffff", Int64.MinusOne
            Test.Use "&H8000000000000000", Int64.MinValue
            Test.Use "&H7fffffffffffffff", Int64.MaxValue
            Test.Use "&H7fffffffffffffff", Int64.MaxValue
            Test.Use "&H07fffffffffffffff", Int64.MaxValue
            Test.Use "&H00007fffffffffffffff", Int64.MaxValue
            Test.Use "&H000007fffffffffffffff", Int64.MaxValue
            
        Case "ParseEx_WithInvalidValueAndAllowHexSpecifier_ThrowsFormatException"
            Test.Use "-1"
            Test.Use " 1"
            Test.Use "1 "
            Test.Use "+1"
            Test.Use "$1"
            Test.Use "1$"
            Test.Use "1 -"
            Test.Use "1e-10"
            Test.Use "1,000"
            Test.Use "1 1"
            Test.Use "1+1"
            Test.Use "1-1"
            Test.Use " ff"
            Test.Use "f f"
            Test.Use "ff "
            Test.Use "1.1"
            Test.Use "1..0"
            Test.Use "++1"
            Test.Use "1++"
            Test.Use "--1"
            Test.Use "1--"
            Test.Use ".0"
            Test.Use " (1) "
            Test.Use "( 1)"
            Test.Use "(1 )"
            Test.Use "1)"
            Test.Use "1e+"
            Test.Use "1e-"
            
        Case "TryParseEx_WithAllowParentheses_ReturnsExpected"
            Test.Use "0", Int64.Zero
            Test.Use "1", Int64.One
            Test.Use "(1)", Int64.MinusOne
        
        Case "ParseEx_WithInvalidValueAndAllowParentheses_ThrowsFormatException"
            Test.Use "-1"
            Test.Use " 1"
            Test.Use "1 "
            Test.Use "+1"
            Test.Use "$1"
            Test.Use "1$"
            Test.Use "1 -"
            Test.Use "1e+10"
            Test.Use "FF"
            Test.Use "1,000"
            Test.Use "1 1"
            Test.Use "1+1"
            Test.Use "1-1"
            Test.Use "1.1"
            Test.Use "1..0"
            Test.Use "++1"
            Test.Use "1++"
            Test.Use "--1"
            Test.Use "1--"
            Test.Use ".0"
            Test.Use " (1) "
            Test.Use "( 1)"
            Test.Use "(1 )"
            Test.Use "1)"
            
        Case "CInt64_WithValues_ReturnsExpected"
            Test.Use 1, Int64.One
            Test.Use -1, Int64.MinusOne
            Test.Use "-1", Int64.MinusOne
            Test.Use CCur(1), Int64.One
            Test.Use CCur(1.5), NewInt64(2, 0)
            Test.Use CCur(-1), Int64.MinusOne
            Test.Use CCur(-1.5), NewInt64(-2, &HFFFFFFFF)
            Test.Use Int64.One, Int64.One
            Test.Use CDec("9223372036854775807"), Int64.MaxValue
            Test.Use CDec("9223372036854775806.6"), Int64.MaxValue
            Test.Use CDec("9223372036854775806.5"), NewInt64(&HFFFFFFFE, &H7FFFFFFF)
            Test.Use CDec("9223372036854775806"), NewInt64(&HFFFFFFFE, &H7FFFFFFF)
            Test.Use CDec("-9223372036854775806"), NewInt64(&H2, &H80000000)
            Test.Use CDec("-9223372036854775808"), Int64.MinValue
            Test.Use CDec("-9223372036854775808.5"), Int64.MinValue
            Test.Use CDec("0"), Int64.Zero
            Test.Use CDec("1"), Int64.One
            Test.Use CDec("-1"), Int64.MinusOne
            Test.Use CDec("-2"), NewInt64(&HFFFFFFFE, &HFFFFFFFF)
            Test.Use 9.22337203685477E+18, CInt64("9223372036854769700")
            Test.Use -9.22337203685477E+18, CInt64("-9223372036854769700")
            Test.Use 0.922337203685477, CInt64(1)
            Test.Use 9.22337203685477E-02, CInt64(0)
            
        Case "CInt64_WithTooLargeValue_ThrowsOverflowException"
            Test.Use CDec("9223372036854775808")
            Test.Use CDec("-9223372036854775809")
            Test.Use 9.22337203685478E+18
            Test.Use 9.22337203685477E+19
            Test.Use "9999999999999999999999999"
            
        Case "Parse_WithMoreThanMaxValue_ThrowsOverflowException"
            Test.Use "9223372036854775808"
            Test.Use "9223372036854775809"
            Test.Use "9223372036854775810"
            Test.Use "9223372036854775817"
            Test.Use "9223372036854775818"
            Test.Use "9223372036854775820"
            Test.Use "&h1ffffffffffffffff"
            Test.Use "9999999999999999999999999"
            
        Case "Max_WithValues_ReturnsExpected"
            Test.Use Int64.One, Int64.MinusOne, Int64.One
            Test.Use Int64.MinusOne, Int64.One, Int64.One
            
        Case "Min_WithValues_ReturnsExpected"
            Test.Use Int64.One, Int64.MinusOne, Int64.MinusOne
            Test.Use Int64.MinusOne, Int64.One, Int64.MinusOne
        
    End Select
End Sub

Public Sub Zero_WhenCalled_ReturnsExpected()
    Dim Actual As Int64
    
    Actual = Int64.Zero
    
    Assert.That Actual, Iz.EqualTo(NewInt64(0, 0))
End Sub

Public Sub One_WhenCalled_ReturnsExpected()
    Dim Actual As Int64
    
    Actual = Int64.One
    
    Assert.That Actual, Iz.EqualTo(NewInt64(1, 0))
End Sub

Public Sub MinusOne_WhenCalled_ReturnsExpected()
    Dim Actual As Int64
    
    Actual = Int64.MinusOne
    
    Assert.That Actual, Iz.EqualTo(NewInt64(-1, -1))
End Sub

Public Sub MinValue_WhenCalled_ReturnsExpected()
    Dim Actual As Int64
    
    Actual = Int64.MinValue
    
    Assert.That Actual, Iz.EqualTo(NewInt64(0, &H80000000))
End Sub

Public Sub MaxValue_WhenCalled_ReturnsExpected()
    Dim Actual As Int64
    
    Actual = Int64.MaxValue
    
    Assert.That Actual, Iz.EqualTo(NewInt64(&HFFFFFFFF, &H7FFFFFFF))
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByRef Left As Int64, ByRef Right As Int64, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Int64.Equals(Left, Right)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Compare_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Int64.Compare(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Sign_WithValues_ReturnsExpected(ByRef a As Int64, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Int64.Sign(a)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Negate_WithValues_ReturnsExpected(ByRef a As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Negate(a)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Absolute_WithValues_ReturnsExpected(ByRef a As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Absolute(a)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Add_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Add(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Subract_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Subtract(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Multiply_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Multiply(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Divide_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Divide(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub DivRem_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Int64.DivRem a, b, Actual
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ShiftLeft_WithValues_ReturnsExpected(ByRef Value As Int64, ByVal BitCount As Long, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.ShiftLeft(Value, BitCount)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ShiftRight_WithValues_ReturnsExpected(ByRef Value As Int64, ByVal BitCount As Long, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.ShiftRight(Value, BitCount)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithValues_ReturnsExpected(ByRef Value As Int64, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Int64.ToString(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFormatting_ReturnsExpected(ByRef Value As Int64, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Int64.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetHashCode_WhenCalled_ShouldGenerateSameCodeThroughObject()
    Dim Value As Int64
    
    Value = NewInt64(5, 12)
    
    Assert.That Int64.GetHashCode(Value), Iz.EqualTo(Object.GetHashCode(Value))
End Sub

Public Sub TryParse_WithValidValues_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParse(s, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TryParseEx_WithAllowTrailingWhiteSpace_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParseEx(s, AllowTrailingWhite, Nothing, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ParseEx_WithInvalidValueAndAllowTrailingWhiteSpace_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    Int64.ParseEx s, AllowTrailingWhite, Nothing
    
    AssertFormatException Err
End Sub

Public Sub TryParseEx_WithAllowCurrencySymbol_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParseEx(s, AllowCurrencySymbol, Nothing, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ParseEx_WithInvalidValueAndAllowCurrencySymbol_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    Int64.ParseEx s, AllowCurrencySymbol, Nothing
    
    AssertFormatException Err
End Sub

Public Sub TryParseEx_WithAllowDecimalPoint_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParseEx(s, AllowDecimalPoint, Nothing, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ParseEx_WithInvalidValueAndAllowDecimalPoint_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    Int64.ParseEx s, AllowDecimalPoint, Nothing
    
    AssertFormatException Err
End Sub

Public Sub TryParseEx_WithAllowExponent_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParseEx(s, AllowExponent, Nothing, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ParseEx_WithInvalidValueAndAllowExponent_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    Int64.ParseEx s, AllowExponent, Nothing
    
    AssertFormatException Err
End Sub

Public Sub TryParseEx_WithAllowHexSpecifier_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParseEx(s, AllowHexSpecifier, Nothing, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TryParse_WithHexPrefix_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParse(s, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ParseEx_WithInvalidValueAndAllowHexSpecifier_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    Int64.ParseEx s, AllowHexSpecifier, Nothing
    
    AssertFormatException Err
End Sub

Public Sub TryParseEx_WithAllowParentheses_ReturnsExpected(ByVal s As String, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Assert.That Int64.TryParseEx(s, AllowParentheses, Nothing, Actual), Iz.True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ParseEx_WithInvalidValueAndAllowParentheses_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    Int64.ParseEx s, AllowParentheses, Nothing
    
    AssertFormatException Err
End Sub

Public Sub CInt64_WithValues_ReturnsExpected(ByRef Value As Variant, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = CInt64(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CInt64_WithTooLargeValue_ThrowsOverflowException(ByRef Value As Variant)
    On Error Resume Next
    
    CInt64 Value
    
    AssertOverflowException Err
End Sub

Public Sub Parse_WithTooManyDigits_ThrowsOverflowException()
    Const M As String = "9223372036854775807"
    On Error Resume Next
    
    Int64.Parse M & "0"
    
    AssertOverflowException Err
End Sub

Public Sub Parse_WithMoreThanMaxValue_ThrowsOverflowException(ByVal s As String)
    On Error Resume Next
    
    Int64.Parse s
    
    AssertOverflowException Err
End Sub

Public Sub Max_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Max(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Min_WithValues_ReturnsExpected(ByRef a As Int64, ByRef b As Int64, ByRef Expected As Int64)
    Dim Actual As Int64
    
    Actual = Int64.Min(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
