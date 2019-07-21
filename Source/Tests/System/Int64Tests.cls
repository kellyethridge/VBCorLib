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
        
    End Select
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByRef Left As Int64, ByRef Right As Int64, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Int64.Equals(Left, Right)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
