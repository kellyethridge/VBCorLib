VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectToStringWithSinglesTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToString_WithSinglesUsingFormat_ReturnsExpected"
            ' General Format
            Test.Use 0!, "", "0"
            Test.Use 12345!, "", "12345"
            Test.Use 12345!, "g", "12345"
            Test.Use -12345!, "g", "-12345"
            Test.Use 1.2345!, "g", "1.2345"
            Test.Use 0.012345!, "g", "0.012345"
            Test.Use 1234.568!, "g", "1234.568"
            Test.Use 1.2345E+25!, "g26", "12345000400000000000000000"
            Test.Use 1.2345E-05!, "", "1.2345E-05"
            Test.Use 1.2345E+15!, "", "1.2345E+15"
            Test.Use 1.2345E-05!, "G", "1.2345E-05"
            Test.Use 1.2345E+15!, "G", "1.2345E+15"
            Test.Use 12345!, "g4", "1.235e+04"
            
            ' Fixed Format
            Test.Use 12345!, "f", "12345.00"
            Test.Use 12345!, "f0", "12345"
            Test.Use 123450!, "f10", "123450.0000000000"
            Test.Use 123.45!, "f", "123.45"
            Test.Use 1.2345!, "f", "1.23"
            Test.Use 1.2345!, "f3", "1.235"
            Test.Use 0.12345!, "f", "0.12"
            Test.Use 1.2345E-04!, "f3", "0.000"
            Test.Use 1.2345E-04!, "f7", "0.0001235"
            Test.Use 1.2345E-04!, "f20", "0.00012345000000000000"
            
            ' Scientific Format
            Test.Use 12345!, "e", "1.234500e+004"
            Test.Use -12345!, "e4", "-1.2345e+004"
            Test.Use 12345!, "e3", "1.235e+004"
            Test.Use 12345!, "e0", "1e+004"
            Test.Use 12345!, "E1", "1.2E+004"
            Test.Use 1234500!, "e", "1.234500e+006"
            Test.Use 1234500!, "e3", "1.235e+006"
            Test.Use 0.0012345!, "e", "1.234500e-003"
            Test.Use 0.0012345!, "e3", "1.235e-003"
            Test.Use 1.2345E-20!, "e7", "1.2345000e-020"
            
            ' Number Format
            Test.Use 1.23!, "n", "1.23"
            Test.Use 1234.56!, "n", "1,234.56"
            Test.Use 1234568!, "n", "1,234,568.00"
            Test.Use 1.234568E+07!, "n", "12,345,680.00"
            Test.Use 123.45!, "n0", "123"
            Test.Use 123.4568!, "n8", "123.45680000"
            Test.Use -123.45!, "n0", "-123"
            Test.Use 0.00123!, "n", "0.00"
            Test.Use 0.001235!, "n5", "0.00124"
            Test.Use 123!, "n", "123.00"
            Test.Use 1.231231E+20!, "n5", "123,123,100,000,000,000,000.00000"
            
            ' Currency Format
            Test.Use 1.23!, "c", "$1.23"
            Test.Use 1234.56!, "c", "$1,234.56"
            Test.Use 1234568!, "c", "$1,234,568.00"
            Test.Use 1.234568E+07!, "c", "$12,345,680.00"
            Test.Use 123.45!, "c0", "$123"
            Test.Use 123.4568!, "c8", "$123.45680000"
            Test.Use -123.45!, "c0", "($123)"
            Test.Use 0.00123!, "c", "$0.00"
            Test.Use 0.001235!, "c5", "$0.00124"
            Test.Use 123!, "c", "$123.00"
            Test.Use 1.231231E+20!, "c5", "$123,123,100,000,000,000,000.00000"
            
            ' Percentage Format
            Test.Use 0!, "p", "0.00 %"
            Test.Use 1.23!, "p", "123.00 %"
            Test.Use 1234.56!, "p", "123,456.00 %"
            Test.Use 1234568!, "p", "123,456,800.00 %"
            Test.Use 1.234568E+07!, "p", "1,234,568,000.00 %"
            Test.Use 123.45!, "p0", "12,345 %"
            Test.Use 123.4568!, "p8", "12,345.68000000 %"
            Test.Use -123.45!, "p0", "-12,345 %"
            Test.Use 0.00123!, "p", "0.12 %"
            Test.Use 0.001235!, "p3", "0.124 %"
            Test.Use 123!, "p", "12,300.00 %"
            Test.Use 1.231231E+20!, "p5", "12,312,310,000,000,000,000,000.00000 %"
            
        Case "ToString_WithSinglesUsingDefaultFormat_ReturnsExpected"
            Test.Use 0!, "0"
            Test.Use 0.1!, "0.1"
            Test.Use 1!, "1"
            Test.Use 0.001, "0.001"
            Test.Use 100.001, "100.001"
            Test.Use 100, "100"
            Test.Use 1234567!, "1234567"
            Test.Use -1234567!, "-1234567"
            Test.Use 1.401298E-45, "1.401298E-45"
            Test.Use 3.402823E+38, "3.402823E+38"
            Test.Use -1.401298E-45, "-1.401298E-45"
            Test.Use -3.402823E+38, "-3.402823E+38"

    End Select
End Sub

Public Sub ToString_WithSinglesUsingFormat_ReturnsExpected(ByVal Value As Single, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Object.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithSinglesUsingDefaultFormat_ReturnsExpected(ByVal Value As Single, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Object.ToString(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithSingleNotANumber_ReturnsNaN()
    Dim d As Single
    On Error Resume Next
    d = 0! / 0!
    On Error GoTo 0
    
    Dim Actual As String
    Actual = Object.ToString(d)
    
    Assert.That Actual, Iz.EqualTo("NaN")
End Sub

Public Sub ToString_WithSingleAsPositiveInfinity_ReturnsInfinity()
    Dim d As Single
    On Error Resume Next
    d = 1! / 0!
    On Error GoTo 0
    
    Dim Actual As String
    Actual = Object.ToString(d)
    
    Assert.That Actual, Iz.EqualTo("Infinity")
End Sub

Public Sub ToString_WithSingleAsNegativeInfinity_ReturnsNegativeInfinity()
    Dim d As Single
    On Error Resume Next
    d = -1! / 0!
    On Error GoTo 0
    
    Dim Actual As String
    Actual = Object.ToString(d)
    
    Assert.That Actual, Iz.EqualTo("-Infinity")
End Sub

