VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConvertTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ConvertTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToDateTime_WithValues_ReturnsExpected"
            Test.Use CorDateTime.MaxValue, CorDateTime.MaxValue
            Test.Use #1/1/2001#, NewDateTime(2001, 1, 1)
            Test.Use Nothing, CorDateTime.MinValue
            Test.Use "1/1/2001", NewDateTime(2001, 1, 1)
            
    End Select
End Sub

Public Sub ToDateTime_WithValues_ReturnsExpected(ByRef Value As Variant, ByVal Expected As CorDateTime)
    Dim Actual As CorDateTime
    
    Set Actual = Convert.ToDateTime(Value)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub ToDateTime_WithInvalidObjectInstanceType_ThrowsInvalidCastException()
    On Error Resume Next
    
    Convert.ToDateTime Me
    
    AssertInvalidCastException Err
End Sub

Public Sub ToDateTime_WithUnsupportedDataType_ThrowsInvalidCastException()
    On Error Resume Next
    
    Convert.ToDateTime 1
    
    AssertInvalidCastException Err
End Sub
