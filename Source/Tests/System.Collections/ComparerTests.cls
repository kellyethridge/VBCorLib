VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ComparerTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private mComparer As Comparer

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Dim Col As New Collection
    
    Select Case Test.MethodName
        Case "LessThan_WithValues_ReturnsExpected"
            Test.Use 1, 2, True
            Test.Use 2, 1, False
            Test.Use 2, 2, False
            Test.Use 1.1, 2.2, True
            Test.Use 2#, 1#, False
            Test.Use "a", "a", False
            Test.Use "a", "b", True
            Test.Use "b", "a", False
            Test.Use Empty, Empty, False
            Test.Use Null, Null, False
            Test.Use Null, 1, True
            Test.Use 1, Null, False
            Test.Use Empty, 1, True
            Test.Use 1, Empty, False
            Test.Use Null, Empty, True
            Test.Use Empty, Null, False
            Test.Use NewInt32(1), NewInt32(2), True
            Test.Use NewInt32(2), NewInt32(1), False
            Test.Use NewInt32(2), NewInt32(2), False
            Test.Use Col, Col, False
            Test.Use Int64.MinValue, Int64.Zero, True
            Test.Use Int64.MaxValue, Int64.Zero, False
            Test.Use Nothing, New Int32, True
            Test.Use New Int32, Nothing, False
            
        Case "GreaterThan_WithValues_ReturnsExpected"
            Test.Use 1, 2, False
            Test.Use 2, 1, True
            Test.Use 2, 2, False
            Test.Use "a", "a", False
            Test.Use "a", "b", False
            Test.Use "b", "a", True
            Test.Use Null, Null, False
            Test.Use Null, 1, False
            Test.Use 1, Null, True
            Test.Use Empty, Empty, False
            Test.Use Empty, 1, False
            Test.Use 1, Empty, True
            Test.Use Null, Empty, False
            Test.Use Empty, Null, True
            Test.Use 1.1, 2.2, False
            Test.Use 2#, 1#, True
            Test.Use NewInt32(1), NewInt32(2), False
            Test.Use NewInt32(2), NewInt32(1), True
            Test.Use NewInt32(2), NewInt32(2), False
            Test.Use Col, Col, False
            Test.Use Int64.MinValue, Int64.Zero, False
            Test.Use Int64.MaxValue, Int64.Zero, True
            Test.Use New Int32, Nothing, True
            Test.Use Nothing, New Int32, False
            
        Case "Equal_WithValues_ReturnsExpected"
            Test.Use 1, 2, False
            Test.Use 2, 1, False
            Test.Use 2, 2, True
            Test.Use "a", "a", True
            Test.Use "a", "b", False
            Test.Use "b", "a", False
            Test.Use Null, Null, True
            Test.Use Empty, Empty, True
            Test.Use Null, 1, False
            Test.Use 1, Null, False
            Test.Use Empty, 1, False
            Test.Use 1, Empty, False
            Test.Use Null, Empty, False
            Test.Use Empty, Null, False
            Test.Use 1.1, 2.2, False
            Test.Use 2#, 1#, False
            Test.Use 2#, 2#, True
            Test.Use NewInt32(1), NewInt32(2), False
            Test.Use NewInt32(2), NewInt32(1), False
            Test.Use NewInt32(2), NewInt32(2), True
            Test.Use Col, Col, True
            Test.Use Int64.MinValue, Int64.Zero, False
            Test.Use Int64.MaxValue, Int64.Zero, False
            Test.Use Int64.MaxValue, Int64.MaxValue, True
            Test.Use Nothing, Nothing, True
            
        Case "Compare_WithDifferentDataTypes_ThrowsArgumentException"
            Test.Use CLng(1), CInt(1)
            Test.Use CLng(1), CByte(1)
            
    End Select
End Sub

Public Sub Setup()
    Set mComparer = Comparer.Default
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

Public Sub Compare_WithNonIComparableObjects_ThrowsArgumentException()
    Dim x As New MockObject
    Dim y As New MockObject
    
    On Error Resume Next
    Comparer.Default.Compare x, y
    
    AssertArgumentException Err
End Sub


