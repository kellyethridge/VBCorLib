VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToStringWithLongsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToString_WithLongsUsingFormat_ReturnsExpected"
            ' Unspecified Format
            Test.Use 0, "", "0"
            Test.Use 1, "", "1"
            Test.Use -1, "", "-1"
            Test.Use 100, "", "100"
            Test.Use &H7FFFFFFF, "", "2147483647"
            Test.Use &H80000000, "", "-2147483648"

            ' General Format
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
        
            ' Currency Format
            Test.Use 1, "c", "$1.00"
            Test.Use 1, "c0", "$1"
            Test.Use 1, "c1", "$1.0"
            Test.Use 1, "c3", "$1.000"
            Test.Use -1, "c", "($1.00)"
            Test.Use 12345, "c", "$12,345.00"
            
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
            Test.Use 255, "x", "ff"
            Test.Use 255, "X", "FF"
            Test.Use 255, "X4", "00FF"
            Test.Use -1, "x", "ffffffff"
            
        Case "ToString_WithLongsUsingDefaultFormat_ReturnsExpected"
            Test.Use 0&, "0"
            Test.Use 1&, "1"
            Test.Use -1&, "-1"
            Test.Use &H7FFFFFFF, "2147483647"
            Test.Use &H80000000, "-2147483648"
    End Select
End Sub

Public Sub ToString_WithLongsUsingFormat_ReturnsExpected(ByVal Value As Long, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Convert.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithLongsUsingDefaultFormat_ReturnsExpected(ByVal Value As Long, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Convert.ToString(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithLongUsingModifiedGroupSize_ReturnsExpected()
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    info.CurrencyGroupSizes = NewLongs(1, 2, 3)
    
    Actual = Convert.ToString(12345678, "c", info)
    
    Assert.That Actual, Iz.EqualTo("$12,345,67,8.00")
End Sub

Public Sub ToString_WithLongUsingModifiedGroupSize_ReturnsExpected2()
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    info.NumberGroupSizes = NewLongs(2, 3)
    
    Actual = Convert.ToString(1, "n", info)
    
    Assert.That Actual, Iz.EqualTo("1.00")
End Sub

