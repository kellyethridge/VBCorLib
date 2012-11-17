VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArraySortTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As simplyvbcomp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Sort_WithPrimitives_SortsAscending"
            Test.Use NewBytes(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewIntegers(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewLongs(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewSingles(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewDoubles(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewCurrencys(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewStrings("c", "d", "a", "B", "A")
            Test.Use NewDates(#5/1/2001#, #4/1/2001#, #6/1/2001#, #3/1/2001#)
            Test.Use NewVariants(7, 8, 9, 6, 5, 4, 1, 3, 2)
            
    End Select
End Sub

Public Sub Sort_WithPrimitives_SortsAscending(ByRef Arr As Variant)
    cArray.Sort Arr
    
    Assert.That Arr, Iz.Ordered
End Sub

Public Sub Sort_WithNullArray_Throws()
    On Error GoTo ErrTrap
    Dim Arr() As Long
    
    cArray.Sort Arr
    
ErrTrap:
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub Sort_WithMultiDimensionArray_Throws()
    On Error GoTo ErrTrap
    Dim Arr() As Long
    ReDim Arr(1, 1)
    
    cArray.Sort Arr
    
ErrTrap:
    AssertRankException Err
End Sub

Public Sub Sort_WithNonComparableValues_Throws()
    On Error GoTo ErrTrap
    Dim Arr As Variant
    Arr = Array(3, 2, "4", 1)
    
    cArray.Sort Arr
    
ErrTrap:
    AssertInvalidOperationException Err
End Sub

Public Sub Sort_WithInvalidComparer_Throws()
    On Error GoTo ErrTrap
    Dim Arr(0) As Long
    
    cArray.Sort Arr, "Comparer"
    
ErrTrap:
    AssertArgumentException Err, "Comparer"
End Sub

Public Sub Sort_WithComparer_SortsValues()
    Dim Arr() As Long
    Arr = NewLongs(7, 8, 9, 6, 5, 4, 1, 3, 2)
    
    cArray.Sort Arr, New DescendingComparer
    
    Assert.That Arr, Iz.Ordered.Descending
End Sub

Public Sub Sort_WithComparableObjects_SortsObjects()
    Dim a(4) As Int32
    Set a(0) = MakeInt32(3)
    Set a(1) = MakeInt32(2)
    Set a(2) = MakeInt32(5)
    Set a(3) = MakeInt32(4)
    Set a(4) = MakeInt32(0)
    
    cArray.Sort a
    
    Assert.That a, Iz.Ordered.Using(New ComparableComparer)
End Sub

Public Sub Sort_WithCallback_SortsStrings()
    Dim a() As String
    a = NewStrings("c", "d", "a", "B", "A")
    
    cArray.Sort a, AddressOf CompareStringsAscending
    
    Assert.That a, Iz.Ordered
End Sub

Public Sub Sort_WithCallback_SortsUDTs()
    Dim a(4) As VBGUID
    a(0).Data1 = 3
    a(1).Data1 = 2
    a(2).Data1 = 4
    a(3).Data1 = 1
    a(4).Data1 = 6
    
    cArray.Sort a, AddressOf CompareVBGuids
    
    Assert.That a, Iz.Ordered.Using(New VBGuidComparer)
End Sub


Public Sub Sort_WithNonZeroLBound_SortsValues()
    Dim Arr(5 To 9) As Long
    Arr(5) = 5
    Arr(6) = 2
    Arr(7) = 3
    Arr(8) = 6
    Arr(9) = 1
    
    cArray.Sort Arr
    
    Assert.That Arr, Iz.Ordered
End Sub

Public Sub SortEx_WithNullArray_Throws()
    On Error GoTo ErrTrap
    Dim a() As Long
    
    cArray.SortEx a, 2, 4
    
ErrTrap:
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub SortEx_WithMultiDimensionArray_Throws()
    On Error GoTo ErrTrap
    Dim Arr(1, 1) As Long
    
    cArray.SortEx Arr, 0, 1
    
ErrTrap:
    AssertRankException Err
End Sub

Public Sub SortEx_WithIndexLessThanLBound_Throws()
    On Error GoTo ErrTrap
    Dim Arr(0) As Long
    
    cArray.SortEx Arr, -1, 1
    
ErrTrap:
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub SortEx_WithLengthLessThanZero_Throws()
    On Error GoTo ErrTrap
    Dim Arr(0) As Long
    
    cArray.SortEx Arr, 0, -1
    
ErrTrap:
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub SortEx_WithRangeExtendingPassedEndOfArray_Throws()
    On Error GoTo ErrTrap
    Dim Arr(1 To 10) As Long
    
    cArray.SortEx Arr, 1, 11
    
ErrTrap:
    AssertArgumentException Err
End Sub

Public Sub SortEx_WhenCalled_SortsSubArray()
    Dim a() As Long
    a = NewLongs(6, 5, 4, 1, 3, 2, 1, 0)
    
    cArray.SortEx a, 2, 4
    
    Dim Expected() As Long
    Expected = NewLongs(6, 5, 1, 2, 3, 4, 1, 0)
    Assert.That a, Iz.EqualTo(Expected)
End Sub

Public Sub SortEx_WithComparer_SortsSubArray()
    Dim a() As Long
    a = NewLongs(5, 6, 4, 1, 3, 2, 1, 0)
    
    cArray.SortEx a, 2, 4, New DescendingComparer
    
    Dim Expected() As Long
    Expected = NewLongs(5, 6, 4, 3, 2, 1, 1, 0)
    Assert.That a, Iz.EqualTo(Expected)
End Sub

Public Sub SortKey_WithNullKeysArray_Throws()
    On Error GoTo ErrTrap
    Dim Keys() As Long
    Dim Items(0) As Long
    
    cArray.SortKey Keys, Items
    
ErrTrap:
    AssertArgumentNullException Err, "Keys"
End Sub

Public Sub SortKey_WithNullItemsArray_Throws()
    On Error GoTo ErrTrap
    Dim Keys(0) As Long
    Dim Items() As Long
    
    cArray.SortKey Keys, Items
    
ErrTrap:
    AssertArgumentNullException Err, "Items"
End Sub

Public Sub SortKey_WithMultiDimensionKeys_Throws()
    On Error GoTo ErrTrap
    Dim Keys(0, 0) As Long
    Dim Items(0) As Long
    
    cArray.SortKey Keys, Items
    
ErrTrap:
    AssertRankException Err
End Sub

Public Sub SortKey_WithMultiDimensionItems_Throws()
    On Error GoTo ErrTrap
    Dim Keys(0) As Long
    Dim Items(0, 0) As Long
    
    cArray.SortKey Keys, Items
    
ErrTrap:
    AssertRankException Err
End Sub

Public Sub SortKey_WithInvalidItems_Throws()
    On Error GoTo ErrTrap
    Dim Keys(0) As Long
    Dim InvalidItems As Long
    
    cArray.SortKey Keys, InvalidItems
    
ErrTrap:
    AssertArgumentException Err, "Items"
End Sub

Public Sub SortKeyEx_WithNullKeys_Throws()
    On Error GoTo ErrTrap
    Dim Keys() As Long
    Dim Items(0) As Long
    
    cArray.SortKeyEx Keys, Items, 0, 1
    
ErrTrap:
    AssertArgumentNullException Err, "Keys"
End Sub

Public Sub SortKey_WhenCalled_SortsKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    cArray.SortKey Keys, Items
    
    Assert.That Keys, Iz.Ordered
End Sub

Public Sub SortKey_WhenCalled_SortsItemsByKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    cArray.SortKey Keys, Items
    
    Dim Expected() As Long
    Expected = NewLongs(5, 3, 1, 2, 4)
    Assert.That Items, Iz.EqualTo(Expected)
End Sub

Public Sub SortKey_WithComparer_SortsKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    cArray.SortKey Keys, Items, New DescendingComparer
    
    Assert.That Keys, Iz.Ordered.Descending
End Sub

Public Sub SortKey_WithComparer_SortsItemsByKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    cArray.SortKey Keys, Items, New DescendingComparer
    
    Dim Expected() As Long
    Expected = NewLongs(4, 2, 1, 3, 5)
    Assert.That Items, Iz.EqualTo(Expected)
End Sub

Private Function MakeInt32(ByVal value As Long) As Int32
    Set MakeInt32 = New Int32
    MakeInt32.mValue = value
End Function



