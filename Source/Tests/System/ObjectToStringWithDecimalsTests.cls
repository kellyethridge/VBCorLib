VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectToStringWithDecimalsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ObjectToStringWithDecimalsTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToString_WithFormat_ReturnsExpected"
            ' Unspecified Format
            Test.Use CDec(0), "", "0"
            Test.Use CDec(1), "", "1"
            Test.Use CDec(0.1), "", "0.1"
            Test.Use CDec(0.001), "", "0.001"
            Test.Use CDec(100.001), "", "100.001"
            Test.Use CDec(100), "", "100"
            Test.Use CDec("79228162514264337593543950335"), "", "79228162514264337593543950335"
            Test.Use CDec("-79228162514264337593543950335"), "", "-79228162514264337593543950335"
            Test.Use CDec("0.0000000000000000000000000001"), "", "0.0000000000000000000000000001"
            Test.Use CDec("-0.0000000000000000000000000001"), "", "-0.0000000000000000000000000001"
            
            ' General Format
            Test.Use 0, "g", "0"
            Test.Use 0, "g1", "0"
            Test.Use 0, "g5", "0"
            Test.Use 1, "g", "1"
            Test.Use 1, "g0", "1"
            Test.Use 1, "g1", "1"
            Test.Use 1, "g2", "1"
            Test.Use 0.1, "g", "0.1"
            Test.Use 0.1, "g0", "0.1"
            Test.Use 0.1, "g1", "0.1"
            Test.Use 0.1, "g2", "0.1"
            Test.Use 10, "g", "10"
            Test.Use 10, "g0", "10"
            Test.Use 10, "g1", "1e+01"
            Test.Use 10, "g2", "10"
            Test.Use 123, "g", "123"
            Test.Use 123, "g0", "123"
            Test.Use 123, "g1", "1e+02"
            Test.Use 123, "g2", "1.2e+02"
            Test.Use 123, "g3", "123"
            Test.Use 123000, "g", "123000"
            Test.Use 123000, "g0", "123000"
            Test.Use 123000, "g1", "1e+05"
            Test.Use 123000, "g2", "1.2e+05"
            Test.Use 123000, "g3", "1.23e+05"
            Test.Use 123000, "g4", "1.23e+05"
            Test.Use 0.000123, "g", "0.000123"
            Test.Use 0.000123, "g0", "0.000123"
            Test.Use 0.000123, "g1", "0.0001"
            Test.Use 0.000123, "g2", "0.00012"
            Test.Use 0.000123, "g3", "0.000123"
            Test.Use 0.000123, "g4", "0.000123"
            Test.Use 100.001, "g", "100.001"
            Test.Use 100.001, "g0", "100.001"
            Test.Use 100.001, "g1", "1e+02"
            Test.Use 100.001, "g2", "1e+02"
            Test.Use 100.001, "g3", "100"
            Test.Use 100.001, "g4", "100"
            Test.Use 100.001, "g5", "100"
            Test.Use 100.001, "g6", "100.001"
            Test.Use 1.2345E+27, "g28", "1234500000000000000000000000"
            Test.Use 12345, "g4", "1.235e+04"
            
            ' Currency Format
            Test.Use 0, "c", "$0.00"
            Test.Use 0, "c0", "$0"
            Test.Use 0, "c1", "$0.0"
            Test.Use 0, "c2", "$0.00"
            Test.Use 0, "c3", "$0.000"
            Test.Use 1000, "c", "$1,000.00"
            Test.Use 1000, "c0", "$1,000"
            Test.Use 1000, "c1", "$1,000.0"
            Test.Use 1000, "c2", "$1,000.00"
            Test.Use 1000, "c3", "$1,000.000"
            Test.Use 1.23, "c", "$1.23"
            Test.Use 1234.56, "c", "$1,234.56"
            Test.Use 1234567.89, "c", "$1,234,567.89"
            Test.Use 12345678.91, "c", "$12,345,678.91"
            Test.Use 123.45, "c0", "$123"
            Test.Use 123.45678, "c8", "$123.45678000"
            Test.Use -123.45, "c0", "($123)"
            Test.Use 0.00123, "c", "$0.00"
            Test.Use 0.001235, "c5", "$0.00124"
            Test.Use 123#, "c", "$123.00"
            Test.Use 1.23123123E+20, "c5", "$123,123,123,000,000,000,000.00000"
            
            ' Scientific Format
            Test.Use CDec(12345), "e", "1.234500e+004"
            Test.Use CDec(-12345), "e4", "-1.2345e+004"
            Test.Use CDec(12345), "e3", "1.235e+004"
            Test.Use CDec(12345), "e0", "1e+004"
            Test.Use CDec(12345), "E1", "1.2E+004"
            Test.Use CDec(1234500), "e", "1.234500e+006"
            Test.Use CDec(1234500), "e3", "1.235e+006"
            Test.Use CDec("0.0012345"), "e3", "1.235e-003"
            Test.Use CDec("0.0012345"), "e", "1.234500e-003"
            Test.Use CDec(1), "e", "1.000000e+000"
            Test.Use CDec(1), "e1", "1.0e+000"
            Test.Use CDec(0), "e", "0.000000e+000"
            Test.Use CDec("0.1234"), "e3", "1.234e-001"
            
            ' Fixed Format
            Test.Use 12345, "f", "12345.00"
            Test.Use 12345, "f0", "12345"
            Test.Use 123450, "f10", "123450.0000000000"
            Test.Use 123.45, "f", "123.45"
            Test.Use 1.2345, "f", "1.23"
            Test.Use 1.2345, "f3", "1.235"
            Test.Use 0.12345, "f", "0.12"
            Test.Use 0.00012345, "f3", "0.000"
            Test.Use 0.00012345, "f7", "0.0001235"
            Test.Use 0.00012345, "f20", "0.00012345000000000000"
            Test.Use 1, "f", "1.00"
            Test.Use 0, "f", "0.00"
            
            ' Number Format
            Test.Use 1.23, "n", "1.23"
            Test.Use 1234.56, "n", "1,234.56"
            Test.Use 1234567.89, "n", "1,234,567.89"
            Test.Use 12345678.91, "n", "12,345,678.91"
            Test.Use 123.45, "n0", "123"
            Test.Use 123.45678, "n8", "123.45678000"
            Test.Use -123.45, "n0", "-123"
            Test.Use 0.00123, "n", "0.00"
            Test.Use 0.001235, "n5", "0.00124"
            Test.Use 123, "n", "123.00"
            Test.Use 1.23123123E+20, "n5", "123,123,123,000,000,000,000.00000"
            Test.Use 0, "n", "0.00"
            
            ' Percent Format
            Test.Use 0, "p", "0.00 %"
            Test.Use 1.23, "p", "123.00 %"
            Test.Use 1234.56, "p", "123,456.00 %"
            Test.Use 1234567.89, "p", "123,456,789.00 %"
            Test.Use 12345678.91, "p", "1,234,567,891.00 %"
            Test.Use 123.45, "p0", "12,345 %"
            Test.Use 123.45678, "p8", "12,345.67800000 %"
            Test.Use -123.45, "p0", "-12,345 %"
            Test.Use 0.00123, "p", "0.12 %"
            Test.Use 0.001235, "p3", "0.124 %"
            Test.Use 123, "p", "12,300.00 %"
            Test.Use 1.23123123E+20, "p5", "12,312,312,300,000,000,000,000.00000 %"
            
            ' Custom Format
            Test.Use CDec(1), "aa2", "aa2"
            
        Case "ToString_WithHexSpecifier_ThrowsFormatException"
            Test.Use "x"
            Test.Use "X"
            
    End Select
End Sub

Public Sub ToString_WithFormat_ReturnsExpected(ByVal Value As Variant, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Object.ToString(CDec(Value), Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithHexSpecifier_ThrowsFormatException(ByVal Format As String)
    On Error GoTo Catch
    
    Object.ToString CDec(0), Format
    
Catch:
    AssertFormatException Err
End Sub
