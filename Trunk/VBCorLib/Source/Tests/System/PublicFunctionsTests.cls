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
            Test.Use "te\txt", "te" & vbTab & "xt"
            Test.Use "te\rxt", "te" & vbCr & "xt"
            Test.Use "te\nxt", "te" & vbLf & "xt"
            
    End Select
End Sub

Public Sub T_WithValue_ReturnsExpected(ByVal Text As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = t(Text)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
