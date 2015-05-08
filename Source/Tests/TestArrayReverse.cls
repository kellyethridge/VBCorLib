VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayReverseTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Reverse_WithValues_ReversesOrder"
            Test.Use NewBytes(4, 3, 2, 1)
            Test.Use NewIntegers(4, 3, 2, 1)
            Test.Use NewLongs(4, 3, 2, 1)
            Test.Use NewStrings("4", "3", "2", "1")
            Test.Use NewDoubles(4, 3, 2, 1)
            Test.Use NewVariants(4, 3, 2, 1)
            
    End Select
End Sub

Public Sub Reverse_WithValues_ReversesOrder(ByRef Arr As Variant)
    CorArray.Reverse Arr
    
    Assert.That Arr, Iz.Ordered
End Sub

Public Sub Reverse_WithNullArray_Throws()
    Dim Arr() As Long
    
    On Error GoTo errTrap
    CorArray.Reverse Arr
    
errTrap:
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub Reverse_WithMultiDimensionArray_Throws()
    Dim Arr(0, 0) As Long
    
    On Error GoTo errTrap
    CorArray.Reverse Arr
    
errTrap:
    AssertRankException Err
End Sub

Public Sub Reverse_WithIndexLessThanLBound_Throws()
    Dim Arr(0) As Long
    
    On Error GoTo errTrap
    CorArray.Reverse Arr, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Reverse_WithIndexGreaterThanUBound_Throws()
    Dim Arr(0) As Long
    
    On Error GoTo errTrap
    CorArray.Reverse Arr, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Reverse_WithIndexAndCountGreaterThanUBound_Throws()
    Dim Arr(0) As Long
    
    On Error GoTo errTrap
    CorArray.Reverse Arr, 0, 2
    
errTrap:
    AssertArgumentException Err
End Sub

Public Sub Reverse_WithIndexAndCount_ReverseSubArray()
    Dim Arr() As Long
    Arr = NewLongs(1, 1, 1, 4, 3, 2, 0, 0, 0)
    
    CorArray.Reverse Arr, 3, 3
    
    Assert.That Arr, Iz.EqualTo(NewLongs(1, 1, 1, 2, 3, 4, 0, 0, 0))
End Sub

Public Sub Reverse_WithUDTs_ReversesOrder()
    Dim Arr(3) As VBGUID
    Arr(0).Data1 = 0
    Arr(1).Data1 = 1
    Arr(2).Data1 = 2
    Arr(3).Data1 = 3
    
    CorArray.Reverse Arr
    
    Assert.That Arr(0).Data1, Iz.EqualTo(3)
    Assert.That Arr(1).Data1, Iz.EqualTo(2)
    Assert.That Arr(2).Data1, Iz.EqualTo(1)
    Assert.That Arr(3).Data1, Iz.EqualTo(0)
End Sub

