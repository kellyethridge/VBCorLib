VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicFunctionsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   PublicFunctionsTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "T_WithValue_ReturnsExpected"
            Test.Use "Text", "Text"
            Test.Use "\t", vbTab
            Test.Use "\r", vbCr
            Test.Use "\n", vbLf
            Test.Use "\a", Chr$(7)
            Test.Use "\b", vbBack
            Test.Use "\f", Chr$(12)
            Test.Use "\v", vbVerticalTab
            Test.Use "\\", "\"
            Test.Use "\0", vbNullChar
            Test.Use "\xd", vbCr
            Test.Use "\x0d", vbCr
            Test.Use "\x00d", vbCr
            Test.Use "\x000D", vbCr
            Test.Use "\x000Dd", vbCr & "d"
            Test.Use "\xdg", vbCr & "g"
            Test.Use "\ud", vbCr
            Test.Use "\u0d", vbCr
            Test.Use "\u00d", vbCr
            Test.Use "\u000D", vbCr
            Test.Use "\u000Dd", vbCr & "d"
            Test.Use "\udg", vbCr & "g"
            Test.Use "\u8A9E", ChrW$(-30050)
        
        Case "T_WithInvalidEscapeSequence_ThrowsFormatException"
            Test.Use "\"
            Test.Use "\x"
            Test.Use "\u"
            Test.Use "\xj"
            Test.Use "\uj"
            
    End Select
End Sub

Public Sub T_WithValue_ReturnsExpected(ByVal Text As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = t(Text)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub T_WithInvalidEscapeSequence_ThrowsFormatException(ByVal Text As String)
    On Error Resume Next
    
    t Text
    
    AssertFormatException Err
End Sub
