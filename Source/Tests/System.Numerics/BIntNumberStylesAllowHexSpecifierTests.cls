VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIntNumberStylesAllowHexSpecifierTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BIntNumberStylesAllowHexSpecifierTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Numerics"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Parse_WithValidValue_ReturnsBigInteger"
            Test.Use "0", NewBytes(0), 0
            Test.Use "00000000", NewBytes(0), 0
            Test.Use "1", NewBytes(1), 1
            Test.Use "7f", NewBytes(&H7F), 1
            Test.Use "100", NewBytes(0, 1), 1
            Test.Use "f", NewBytes(&HFF), -1
            Test.Use "ff", NewBytes(&HFF), -1
            Test.Use "8", NewBytes(&HF8), -1
            Test.Use "80", NewBytes(&H80), -1
            Test.Use "800", NewBytes(0, &HF8), -1
            Test.Use "8000", NewBytes(0, &H80), -1
            Test.Use "80000", NewBytes(0, 0, &HF8), -1
            Test.Use "10000", NewBytes(0, 0, 1), 1
            
        Case "Parse_WithInvalidValue_ThrowsFormatException"
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
        
        Case "Parse_WithHexPrefix_ReturnsExpected"
            Test.Use "&h0", NewBytes(0), 0
            Test.Use "&h00000000", NewBytes(0), 0
            Test.Use "&H1", NewBytes(1), 1
            Test.Use "&h7f", NewBytes(&H7F), 1
            Test.Use "&h100", NewBytes(0, 1), 1
            Test.Use "0xf", NewBytes(&HFF), -1
            Test.Use "&hff", NewBytes(&HFF), -1
            Test.Use "&h8", NewBytes(&HF8), -1
            Test.Use "0X80", NewBytes(&H80), -1
            Test.Use "&h800", NewBytes(0, &HF8), -1
            Test.Use "&h8000", NewBytes(0, &H80), -1
            Test.Use "&h80000", NewBytes(0, 0, &HF8), -1
            Test.Use "&h10000", NewBytes(0, 0, 1), 1
    End Select
End Sub

Public Sub Parse_WithValidValue_ReturnsBigInteger(ByVal s As String, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Parse(s, NumberStyles.AllowHexSpecifier)
    
    AssertNumber Actual, ExpectedBytes, ExpectedSign
End Sub

Public Sub Parse_WithInvalidValue_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    BigInteger.Parse s, NumberStyles.AllowHexSpecifier
    
    AssertFormatException Err
End Sub

Public Sub Parse_WithHexPrefix_ReturnsExpected(ByVal s As String, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Parse(s)
    
    AssertNumber Actual, ExpectedBytes, ExpectedSign
End Sub


