VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDefaultComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private comparer As IComparer

Public Sub Setup()
    Set comparer = VBCorLib.comparer.Default
End Sub
Public Sub TestCompareNumbers()
    Assert.AreEqual 0, comparer.Compare(1, 1), "Should return 0."
    Assert.IsTrue comparer.Compare(1, 2) < 0, "Should return negative number."
    Assert.IsTrue comparer.Compare(2, 1) > 0, "Should return positive number."
End Sub

Public Sub TestCompareFloats()
    Assert.AreEqual 0, comparer.Compare(2.5, 2.5), "Should return 0."
    Assert.IsTrue comparer.Compare(1.25, 2.5) < 0, "Should return negative number."
    Assert.IsTrue comparer.Compare(2.5, 1.25) > 0, "Should return positive number."
End Sub

Public Sub TestCompareNulls()
    Assert.IsTrue comparer.Compare(Null, 0) < 0, "Should return negative number."
    Assert.IsTrue comparer.Compare(0, Null) > 0, "Should return positive number."
    Assert.AreEqual 0, comparer.Compare(Null, Null), "Should return 0."
End Sub

Public Sub TestCompareEmpty()
    Assert.AreEqual 0, comparer.Compare(Empty, Empty), "Should return 0."
    Assert.IsTrue comparer.Compare(Empty, 0) < 0, "Should return negative number."
    Assert.IsTrue comparer.Compare(0, Empty) > 0, "Should return positive number."
End Sub
