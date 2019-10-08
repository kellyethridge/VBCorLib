VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CaseInsensitiveComparerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CaseInsensitiveComparerTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private mComparer As CaseInsensitiveComparer

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "LessThan_WithValues_ReturnsExpected"
            Test.Use 1, 2, True
            Test.Use 2, 1, False
            Test.Use 2, 2, False
            Test.Use "a", "B", True
            Test.Use "A", "B", True
            Test.Use "A", "a", False
            Test.Use "a", "A", False
            
        Case "GreaterThan_WithValues_ReturnsExpected"
            Test.Use 1, 2, False
            Test.Use 2, 1, True
            Test.Use 2, 2, False
            Test.Use "B", "a", True
            Test.Use "B", "A", True
            Test.Use "A", "a", False
            Test.Use "a", "A", False
            
        Case "Equal_WithValues_ReturnsExpected"
            Test.Use 1, 2, False
            Test.Use 2, 1, False
            Test.Use 2, 2, True
            Test.Use "a", "B", False
            Test.Use "A", "B", False
            Test.Use "A", "a", True
            Test.Use "a", "A", True
            
        Case "Compare_WithDifferentDataTypes_ThrowsArgumentException"
            Test.Use CLng(1), CInt(1)
            Test.Use CLng(1), CByte(1)
            
    End Select
End Sub

Public Sub Setup()
    Set mComparer = CaseInsensitiveComparer.Default
End Sub

Public Sub LessThan_WithValues_ReturnsExpected(ByRef x As Variant, ByRef y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = mComparer.Compare(x, y) < 0
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GreaterThan_WithValues_ReturnsExpected(ByRef x As Variant, ByRef y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = mComparer.Compare(x, y) > 0
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equal_WithValues_ReturnsExpected(ByRef x As Variant, ByRef y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = mComparer.Compare(x, y) = 0
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Compare_WithDifferentDataTypes_ThrowsArgumentException(ByRef x As Variant, ByRef y As Variant)
    On Error Resume Next
    
    mComparer.Compare x, y

    AssertArgumentException Err
End Sub
