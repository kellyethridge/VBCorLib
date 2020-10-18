VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIntNumberStylesAllowExponentTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BIntNumberStylesAllowExponentTests
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
            Test.Use "1E0", NewBytes(1), 1
            Test.Use "1e2", NewBytes(&H64), 1
            Test.Use "1e10", NewBytes(&H0, &HE4, &HB, &H54, &H2), 1
            Test.Use "1e+10", NewBytes(&H0, &HE4, &HB, &H54, &H2), 1
            
        Case "Parse_WithInvalidValue_ThrowsFormatException"
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
            
    End Select
End Sub

Public Sub Parse_WithValidValue_ReturnsBigInteger(ByVal s As String, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Parse(s, NumberStyles.AllowExponent)
    
    AssertNumber Actual, ExpectedBytes, ExpectedSign
End Sub

Public Sub Parse_WithInvalidValue_ThrowsFormatException(ByVal s As String)
    On Error Resume Next
    
    BigInteger.Parse s, NumberStyles.AllowExponent
    
    AssertFormatException Err
End Sub

