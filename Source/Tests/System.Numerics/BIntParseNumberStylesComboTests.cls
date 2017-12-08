VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIntParseNumberStylesComboTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BIntParseNumberStylesComboTests
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
            Test.Use " -1", NumberStyles.Integer, NewBytes(&HFF), -1
            Test.Use "-1 ", NumberStyles.Integer, NewBytes(&HFF), -1
            Test.Use " +1", NumberStyles.Integer, NewBytes(1), 1
            Test.Use "+1 ", NumberStyles.Integer, NewBytes(1), 1
            Test.Use " $1", NumberStyles.Integer Or NumberStyles.AllowCurrencySymbol, NewBytes(1), 1
            Test.Use " $ 1", NumberStyles.Integer Or NumberStyles.AllowCurrencySymbol, NewBytes(1), 1
            Test.Use " 1$ ", NumberStyles.Integer Or NumberStyles.AllowCurrencySymbol, NewBytes(1), 1
            Test.Use " 1 $ ", NumberStyles.Integer Or NumberStyles.AllowCurrencySymbol, NewBytes(1), 1
            Test.Use " -1 $ ", NumberStyles.Integer Or NumberStyles.AllowCurrencySymbol, NewBytes(&HFF), -1
            Test.Use "1,.0", NumberStyles.AllowDecimalPoint Or NumberStyles.AllowThousands, NewBytes(1), 1
            Test.Use " (1) ", NumberStyles.Integer Or NumberStyles.AllowParentheses, NewBytes(&HFF), -1
            Test.Use "(1 )", NumberStyles.Integer Or NumberStyles.AllowParentheses, NewBytes(&HFF), -1
            Test.Use "(1.)", NumberStyles.Integer Or NumberStyles.AllowParentheses Or NumberStyles.AllowDecimalPoint, NewBytes(&HFF), -1
            Test.Use "(.0)", NumberStyles.Integer Or NumberStyles.AllowParentheses Or NumberStyles.AllowDecimalPoint, NewBytes(0), 0
            
        Case "Parse_WithInvalidValue_ThrowsFormatException"
            Test.Use "( 1)", NumberStyles.Integer Or NumberStyles.AllowParentheses
            Test.Use "+(1)", NumberStyles.Integer Or NumberStyles.AllowParentheses
            Test.Use "(+1)", NumberStyles.Integer Or NumberStyles.AllowParentheses
            Test.Use "-(1)", NumberStyles.Integer Or NumberStyles.AllowParentheses
            Test.Use "(-1)", NumberStyles.Integer Or NumberStyles.AllowParentheses
            Test.Use "(1).", NumberStyles.Integer Or NumberStyles.AllowParentheses Or NumberStyles.AllowDecimalPoint
            Test.Use ".(1)", NumberStyles.Integer Or NumberStyles.AllowParentheses Or NumberStyles.AllowDecimalPoint
            
        Case "Parse_WithValidValueAndCustomCurrencyDecimalSeparator_ReturnsBigInteger"
            Test.Use "1.0"
            Test.Use "1[0"
            Test.Use "$1[0"
            
        Case "Parse_WithInvalidValueAndCustomCurrencyDecimalSeparator_ThrowsFormatException"
            Test.Use "$1.0"
            
        Case "Parse_WithValidValueAndCustomCurrencyThousandsSeparator_ReturnsBigInteger"
            Test.Use "1,0"
            Test.Use "1[0"
            Test.Use "$1[0"
            
        Case "Parse_WithInvalidValueAndCustomCurrencyThousandsSeparator_ThrowsFormatException"
            Test.Use "$1,0"
            
    End Select
End Sub

Public Sub Parse_WithValidValue_ReturnsBigInteger(ByVal s As String, ByVal Style As NumberStyles, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Dim Actual As BigInteger
    
    Set Actual = BigInteger.Parse(s, Style)
    
    AssertNumber Actual, ExpectedBytes, ExpectedSign
End Sub

Public Sub Parse_WithInvalidValue_ThrowsFormatException(ByVal s As String, ByVal Style As NumberStyles)
    On Error Resume Next
    
    BigInteger.Parse s, Style
    
    AssertFormatException Err
End Sub

Public Sub Parse_WithValidValueAndCustomCurrencyDecimalSeparator_ReturnsBigInteger(ByVal s As String)
    Dim Info As NumberFormatInfo
    Set Info = NumberFormatInfo.CurrentInfo.Clone
    Info.CurrencyDecimalSeparator = "["
    
    Dim Actual As BigInteger
    Set Actual = BigInteger.Parse(s, NumberStyles.AllowDecimalPoint Or NumberStyles.AllowCurrencySymbol, Info)
    
    AssertNumber Actual, NewBytes(1), 1
End Sub

Public Sub Parse_WithInvalidValueAndCustomCurrencyDecimalSeparator_ThrowsFormatException(ByVal s As String)
    Dim Info As NumberFormatInfo
    Set Info = NumberFormatInfo.CurrentInfo.Clone
    Info.CurrencyDecimalSeparator = "["
    
    On Error Resume Next
    
    Dim Actual As BigInteger
    Set Actual = BigInteger.Parse(s, NumberStyles.AllowDecimalPoint Or NumberStyles.AllowCurrencySymbol, Info)

    AssertFormatException Err
End Sub

Public Sub Parse_WithValidValueAndCustomCurrencyThousandsSeparator_ReturnsBigInteger(ByVal s As String)
    Dim Info As NumberFormatInfo
    Set Info = NumberFormatInfo.CurrentInfo.Clone
    Info.CurrencyGroupSeparator = "["
    
    Dim Actual As BigInteger
    Set Actual = BigInteger.Parse(s, NumberStyles.AllowThousands Or NumberStyles.AllowCurrencySymbol, Info)
    
    AssertNumber Actual, NewBytes(10), 1
End Sub

Public Sub Parse_WithInvalidValueAndCustomCurrencyThousandsSeparator_ThrowsFormatException(ByVal s As String)
    Dim Info As NumberFormatInfo
    Set Info = NumberFormatInfo.CurrentInfo.Clone
    Info.CurrencyGroupSeparator = "["
    
    On Error Resume Next
    
    Dim Actual As BigInteger
    Set Actual = BigInteger.Parse(s, NumberStyles.AllowThousands Or NumberStyles.AllowCurrencySymbol, Info)

    AssertFormatException Err
End Sub
