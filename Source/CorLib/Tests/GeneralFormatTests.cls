VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneralFormatTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

'Public Sub FixtureSetup()
'    Thread.CurrentThread.CurrentCulture = NewCultureInfo("en-US")
'End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToString_WithLongs_ReturnsExpected"
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
            Test.Use 0, "g1", "0"
            Test.Use 1, "g1", "1"
            Test.Use -1, "g1", "-1"
            Test.Use &H7FFFFFFF, "g1", "2e+09"
            Test.Use &H80000000, "g1", "-2e+09"
            Test.Use 0, "g5", "0"
            Test.Use 1, "g5", "1"
            Test.Use -1, "g5", "-1"
            Test.Use &H7FFFFFFF, "g5", "2.1475e+09"
            Test.Use &H80000000, "g5", "-2.1475e+09"
            Test.Use 12345, "ggg", "ggg"
    
        Case "ToString_WithDoubles_ReturnsExpected"
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
            
        Case "ToString_WithSingles_ReturnsExpected"
            Test.Use 12345!, "g", "12345"
            Test.Use -12345!, "g", "-12345"
            Test.Use 1.2345!, "g", "1.2345"
            Test.Use 0.012345!, "g", "0.012345"
            Test.Use 1234.568!, "g", "1234.568"
            Test.Use 1.2345E+25!, "g26", "12345000400000000000000000"
            Test.Use 1.2345E-05!, "G", "1.2345E-05"
            Test.Use 1.2345E+15!, "G", "1.2345E+15"
            Test.Use 12345!, "g4", "1.235e+04"

    End Select
End Sub

Public Sub ToString_WithLongs_ReturnsExpected(ByVal Value As Long, ByVal Format As String, ByVal Expected As String)
    ValidateFormat Format
    Dim Actual As String
    
    Actual = Convert.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithDoubles_ReturnsExpected(ByVal Value As Double, ByVal Format As String, ByVal Expected As String)
    ValidateFormat Format
    Dim Actual As String
    
    Actual = Convert.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithSingles_ReturnsExpected(ByVal Value As Single, ByVal Format As String, ByVal Expected As String)
    ValidateFormat Format
    Dim Actual As String
    
    Actual = Convert.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub



Private Sub ValidateFormat(ByRef Format As String)
    Assert.That Format, Iz.StringStarting("g").IgnoreCase
End Sub

