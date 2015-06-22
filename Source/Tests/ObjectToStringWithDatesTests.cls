VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectToStringWithDatesTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToString_WithDatesUsingDefaultFormat_ReturnsExpected"
            Test.Use #12:00:00 AM#, "12/30/1899 12:00:00 AM"
            Test.Use #1/1/100#, "1/1/0100 12:00:00 AM"
            Test.Use #12/31/9999 11:59:59 PM#, "12/31/9999 11:59:59 PM"
            Test.Use CorDateTime.MinValue, "1/1/0001 12:00:00 AM"
            Test.Use CorDateTime.MaxValue, "12/31/9999 11:59:59 PM"
            
        
            
    End Select
End Sub

Public Sub ToString_WithFormats_ReturnsExpected(ByRef Value As Variant, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Object.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithDefaultFormat_ReturnsExpected(ByRef Value As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Object.ToString(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
