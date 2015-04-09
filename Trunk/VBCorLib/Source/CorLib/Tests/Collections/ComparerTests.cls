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
Option Explicit
Implements ITestCaseSource

Private mComparer As IComparer

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "LessThan_WithValues_ReturnsExpected"
            Test.Use 1, 2, True
            Test.Use 2, 1, False
            Test.Use 2, 2, False
            Test.Use 1.1, 2.2, True
            Test.Use 2#, 1#, False
            Test.Use Empty, Empty, False
            Test.Use Null, Null, False
            Test.Use NewInt32(1), NewInt32(2), True
            Test.Use NewInt32(2), NewInt32(1), False
            Test.Use NewInt32(2), NewInt32(2), False
            
        Case "GreaterThan_WithValues_ReturnsExpected"
            Test.Use 1, 2, False
            Test.Use 2, 1, True
            Test.Use 2, 2, False
            Test.Use Null, Null, False
            Test.Use Empty, Empty, False
            Test.Use 1.1, 2.2, False
            Test.Use 2#, 1#, True
            Test.Use NewInt32(1), NewInt32(2), False
            Test.Use NewInt32(2), NewInt32(1), True
            Test.Use NewInt32(2), NewInt32(2), False
        
        Case "Equal_WithValues_ReturnsExpected"
            Test.Use 1, 2, False
            Test.Use 2, 1, False
            Test.Use 2, 2, True
            Test.Use Null, Null, True
            Test.Use Empty, Empty, True
            Test.Use 1.1, 2.2, False
            Test.Use 2#, 1#, False
            Test.Use 2#, 2#, True
            Test.Use NewInt32(1), NewInt32(2), False
            Test.Use NewInt32(2), NewInt32(1), False
            Test.Use NewInt32(2), NewInt32(2), True
        
        Case "Compare_WithDifferentDataTypes_ThrowsArgumentException"
            Test.Use CLng(1), CInt(1)
            Test.Use CLng(1), CByte(1)
            
    End Select
End Sub

Public Sub Setup()
    Set mComparer = Comparer.Default
End Sub

Public Sub LessThan_WithValues_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = mComparer.Compare(X, Y) < 0
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GreaterThan_WithValues_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = mComparer.Compare(X, Y) > 0
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equal_WithValues_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = mComparer.Compare(X, Y) = 0
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Compare_WithDifferentDataTypes_ThrowsArgumentException(ByRef X As Variant, ByRef Y As Variant)
    On Error Resume Next
    
    mComparer.Compare X, Y
    
    AssertArgumentException Err
End Sub

'Private Comparer As IComparer
'
'Public Sub Setup()
'    Set Comparer = VBCorLib.Comparer.Default
'End Sub
'
'Public Sub TestCompareNumbers()
'    Assert.AreEqual 0, Comparer.Compare(1, 1), "Should return 0."
'    Assert.IsTrue Comparer.Compare(1, 2) < 0, "Should return negative number."
'    Assert.IsTrue Comparer.Compare(2, 1) > 0, "Should return positive number."
'End Sub
'
'Public Sub TestCompareFloats()
'    Assert.AreEqual 0, Comparer.Compare(2.5, 2.5), "Should return 0."
'    Assert.IsTrue Comparer.Compare(1.25, 2.5) < 0, "Should return negative number."
'    Assert.IsTrue Comparer.Compare(2.5, 1.25) > 0, "Should return positive number."
'End Sub
'
'Public Sub TestCompareNulls()
'    Assert.IsTrue Comparer.Compare(Null, 0) < 0, "Should return negative number."
'    Assert.IsTrue Comparer.Compare(0, Null) > 0, "Should return positive number."
'    Assert.AreEqual 0, Comparer.Compare(Null, Null), "Should return 0."
'End Sub
'
'Public Sub TestCompareEmpty()
'    Assert.AreEqual 0, Comparer.Compare(Empty, Empty), "Should return 0."
'    Assert.IsTrue Comparer.Compare(Empty, 0) < 0, "Should return negative number."
'    Assert.IsTrue Comparer.Compare(0, Empty) > 0, "Should return positive number."
'End Sub

